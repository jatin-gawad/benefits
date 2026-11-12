using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using RenewalReviewBuilder.App.Models;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;
using Windows.Storage.Streams;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using WinPdfDocument = Windows.Data.Pdf.PdfDocument;

namespace RenewalReviewBuilder.App.Services.FlexibleBenefits;

public interface IFlexibleBenefitDocumentTextExtractor
{
    Task<FlexibleBenefitOcrDocument> ExtractAsync(string filePath, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<FlexibleBenefitOcrLine>> ExtractRegionAsync(
        string filePath,
        int pageNumber,
        FlexibleBenefitBounds sourceRegion,
        int sourcePixelWidth,
        int sourcePixelHeight,
        CancellationToken cancellationToken = default);
}

public sealed class WindowsOcrPdfExtractor : IFlexibleBenefitDocumentTextExtractor
{
    private const uint PreferredRenderWidth = 1700;

    public async Task<FlexibleBenefitOcrDocument> ExtractAsync(
        string filePath,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            throw new FileNotFoundException("The flexible-benefits PDF could not be found.", filePath);
        }

        var fullPath = Path.GetFullPath(filePath);
        var fileHash = await ComputeHashAsync(fullPath, cancellationToken);
        var storageFile = await StorageFile.GetFileFromPathAsync(fullPath);
        var winDocument = await WinPdfDocument.LoadFromFileAsync(storageFile);
        using var pdfPigDocument = TryOpenPdfPig(fullPath);
        var ocrEngine = OcrEngine.TryCreateFromUserProfileLanguages()
            ?? throw new InvalidOperationException("Windows OCR is not available for the current Windows user profile.");

        var output = new FlexibleBenefitOcrDocument
        {
            FilePath = fullPath,
            FileHash = fileHash,
            ExtractionMethod = "PdfPig text with local Windows Media OCR fallback",
            Confidence = 0.82m
        };

        for (var pageIndex = 0u; pageIndex < winDocument.PageCount; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var page = winDocument.GetPage(pageIndex);
            using var renderedStream = new InMemoryRandomAccessStream();
            var renderOptions = BuildRenderOptions(page);
            await page.RenderToStreamAsync(renderedStream, renderOptions);
            renderedStream.Seek(0);

            var decoder = await BitmapDecoder.CreateAsync(renderedStream);
            using var decodedBitmap = await decoder.GetSoftwareBitmapAsync();
            using var bitmap = SoftwareBitmap.Convert(
                decodedBitmap,
                BitmapPixelFormat.Bgra8,
                BitmapAlphaMode.Ignore);

            var pageNumber = checked((int)pageIndex + 1);
            var pageResult = HasMachineReadableText(pdfPigDocument, pageNumber)
                ? BuildMachineReadablePage(pdfPigDocument!, pageNumber, bitmap.PixelWidth, bitmap.PixelHeight)
                : await BuildOcrPageAsync(ocrEngine, bitmap, pageNumber);

            foreach (var mark in DetectSelectionMarks(bitmap, pageResult.Lines))
            {
                pageResult.Marks.Add(mark);
            }

            output.Pages.Add(pageResult);
        }

        return output;
    }

    public async Task<IReadOnlyList<FlexibleBenefitOcrLine>> ExtractRegionAsync(
        string filePath,
        int pageNumber,
        FlexibleBenefitBounds sourceRegion,
        int sourcePixelWidth,
        int sourcePixelHeight,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            throw new FileNotFoundException("The flexible-benefits PDF could not be found.", filePath);
        }

        if (pageNumber < 1 || sourcePixelWidth <= 0 || sourcePixelHeight <= 0)
        {
            return [];
        }

        cancellationToken.ThrowIfCancellationRequested();
        var storageFile = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(filePath));
        var document = await WinPdfDocument.LoadFromFileAsync(storageFile);
        if (pageNumber > document.PageCount)
        {
            return [];
        }

        using var page = document.GetPage(checked((uint)pageNumber - 1));
        using var renderedStream = new InMemoryRandomAccessStream();
        var mediaBox = page.Dimensions.MediaBox;
        var sourceRect = new Windows.Foundation.Rect(
            mediaBox.X + Math.Max(0d, sourceRegion.X) / sourcePixelWidth * mediaBox.Width,
            mediaBox.Y + Math.Max(0d, sourceRegion.Y) / sourcePixelHeight * mediaBox.Height,
            Math.Min(sourceRegion.Width, sourcePixelWidth - Math.Max(0d, sourceRegion.X)) / sourcePixelWidth * mediaBox.Width,
            Math.Min(sourceRegion.Height, sourcePixelHeight - Math.Max(0d, sourceRegion.Y)) / sourcePixelHeight * mediaBox.Height);
        var maxDimension = Math.Max(1000u, OcrEngine.MaxImageDimension);
        var targetWidth = (uint)Math.Min(maxDimension, Math.Max(700d, sourceRegion.Width * 3d));
        var targetHeight = (uint)Math.Max(1d, Math.Round(targetWidth * sourceRegion.Height / Math.Max(1d, sourceRegion.Width)));
        if (targetHeight > maxDimension)
        {
            targetHeight = maxDimension;
            targetWidth = (uint)Math.Max(1d, Math.Round(targetHeight * sourceRegion.Width / Math.Max(1d, sourceRegion.Height)));
        }

        await page.RenderToStreamAsync(renderedStream, new PdfPageRenderOptions
        {
            SourceRect = sourceRect,
            DestinationWidth = targetWidth,
            DestinationHeight = targetHeight
        });
        renderedStream.Seek(0);
        var decoder = await BitmapDecoder.CreateAsync(renderedStream);
        using var decodedBitmap = await decoder.GetSoftwareBitmapAsync();
        using var bitmap = SoftwareBitmap.Convert(decodedBitmap, BitmapPixelFormat.Bgra8, BitmapAlphaMode.Ignore);
        var scaleX = bitmap.PixelWidth / Math.Max(1d, sourceRegion.Width);
        var scaleY = bitmap.PixelHeight / Math.Max(1d, sourceRegion.Height);
        var engine = OcrEngine.TryCreateFromUserProfileLanguages()
            ?? throw new InvalidOperationException("Windows OCR is not available for the current Windows user profile.");
        var result = await engine.RecognizeAsync(bitmap);
        var output = new List<FlexibleBenefitOcrLine>();
        foreach (var sourceLine in result.Lines)
        {
            var words = sourceLine.Words.Select(word => new FlexibleBenefitOcrWord
            {
                Text = Clean(word.Text),
                Bounds = new FlexibleBenefitBounds(
                    sourceRegion.X + word.BoundingRect.X / scaleX,
                    sourceRegion.Y + word.BoundingRect.Y / scaleY,
                    word.BoundingRect.Width / scaleX,
                    word.BoundingRect.Height / scaleY)
            }).Where(word => word.Text.Length > 0).ToList();
            if (words.Count == 0)
            {
                continue;
            }

            var line = new FlexibleBenefitOcrLine
            {
                Text = Clean(string.Join(" ", words.Select(word => word.Text))),
                Bounds = Union(words.Select(word => word.Bounds))
            };
            line.Words.AddRange(words);
            output.Add(line);
        }

        if (!output.Any(line => Regex.IsMatch(line.Text, @"\d")))
        {
            var isolatedZero = DetectIsolatedZero(bitmap, sourceRegion, scaleX, scaleY, output);
            if (isolatedZero is not null)
            {
                output.Add(isolatedZero);
            }
        }

        return output;
    }

    private static FlexibleBenefitOcrLine? DetectIsolatedZero(
        SoftwareBitmap bitmap,
        FlexibleBenefitBounds sourceRegion,
        double scaleX,
        double scaleY,
        IReadOnlyList<FlexibleBenefitOcrLine> recognizedLines)
    {
        var width = bitmap.PixelWidth;
        var height = bitmap.PixelHeight;
        var buffer = new Windows.Storage.Streams.Buffer(checked((uint)(width * height * 4)));
        bitmap.CopyToBuffer(buffer);
        var pixels = new byte[buffer.Length];
        using (var reader = DataReader.FromBuffer(buffer))
        {
            reader.ReadBytes(pixels);
        }

        var dark = new bool[width * height];
        for (var index = 0; index < dark.Length; index++)
        {
            var offset = index * 4;
            dark[index] = (pixels[offset] + pixels[offset + 1] + pixels[offset + 2]) / 3 < 140;
        }

        var recognizedRight = recognizedLines.Count == 0
            ? 0d
            : recognizedLines.Max(line => (line.Bounds.Right - sourceRegion.X) * scaleX);
        var minimumX = Math.Clamp((int)Math.Floor(recognizedRight + width * 0.01), 0, width - 1);
        var visited = new bool[dark.Length];
        var queue = new Queue<int>();
        var candidates = new List<(int MinX, int MinY, int MaxX, int MaxY, int Count)>();
        for (var y = 0; y < height; y++)
        {
            for (var x = minimumX; x < width; x++)
            {
                var start = y * width + x;
                if (!dark[start] || visited[start])
                {
                    continue;
                }

                visited[start] = true;
                queue.Enqueue(start);
                var minX = width;
                var minY = height;
                var maxX = 0;
                var maxY = 0;
                var count = 0;
                while (queue.Count > 0)
                {
                    var current = queue.Dequeue();
                    var currentX = current % width;
                    var currentY = current / width;
                    minX = Math.Min(minX, currentX);
                    minY = Math.Min(minY, currentY);
                    maxX = Math.Max(maxX, currentX);
                    maxY = Math.Max(maxY, currentY);
                    count++;
                    Visit(currentX - 1, currentY);
                    Visit(currentX + 1, currentY);
                    Visit(currentX, currentY - 1);
                    Visit(currentX, currentY + 1);
                }

                var componentWidth = maxX - minX + 1;
                var componentHeight = maxY - minY + 1;
                var ratio = componentHeight == 0 ? 0d : componentWidth / (double)componentHeight;
                if (componentHeight >= height * 0.10 && componentHeight <= height * 0.70 &&
                    componentWidth >= width * 0.015 && ratio is >= 0.30 and <= 1.20 &&
                    count < componentWidth * componentHeight * 0.70 &&
                    HasEnclosedWhiteArea(dark, width, height, minX, minY, maxX, maxY))
                {
                    candidates.Add((minX, minY, maxX, maxY, count));
                }

                void Visit(int nextX, int nextY)
                {
                    if (nextX < minimumX || nextY < 0 || nextX >= width || nextY >= height)
                    {
                        return;
                    }

                    var next = nextY * width + nextX;
                    if (!dark[next] || visited[next])
                    {
                        return;
                    }

                    visited[next] = true;
                    queue.Enqueue(next);
                }
            }
        }

        var glyph = candidates
            .OrderBy(candidate => candidate.MinX)
            .ThenByDescending(candidate => candidate.Count)
            .FirstOrDefault();
        if (glyph == default)
        {
            return null;
        }

        var bounds = new FlexibleBenefitBounds(
            sourceRegion.X + glyph.MinX / scaleX,
            sourceRegion.Y + glyph.MinY / scaleY,
            (glyph.MaxX - glyph.MinX + 1) / scaleX,
            (glyph.MaxY - glyph.MinY + 1) / scaleY);
        var word = new FlexibleBenefitOcrWord { Text = "0", Bounds = bounds };
        var line = new FlexibleBenefitOcrLine { Text = "0", Bounds = bounds };
        line.Words.Add(word);
        return line;
    }

    private static bool HasEnclosedWhiteArea(
        IReadOnlyList<bool> dark,
        int imageWidth,
        int imageHeight,
        int minX,
        int minY,
        int maxX,
        int maxY)
    {
        minX = Math.Max(0, minX - 1);
        minY = Math.Max(0, minY - 1);
        maxX = Math.Min(imageWidth - 1, maxX + 1);
        maxY = Math.Min(imageHeight - 1, maxY + 1);
        var localWidth = maxX - minX + 1;
        var localHeight = maxY - minY + 1;
        var outside = new bool[localWidth * localHeight];
        var queue = new Queue<int>();
        for (var x = 0; x < localWidth; x++)
        {
            Enqueue(x, 0);
            Enqueue(x, localHeight - 1);
        }

        for (var y = 0; y < localHeight; y++)
        {
            Enqueue(0, y);
            Enqueue(localWidth - 1, y);
        }

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            var x = current % localWidth;
            var y = current / localWidth;
            Enqueue(x - 1, y);
            Enqueue(x + 1, y);
            Enqueue(x, y - 1);
            Enqueue(x, y + 1);
        }

        var enclosedCount = 0;
        for (var y = 1; y < localHeight - 1; y++)
        {
            for (var x = 1; x < localWidth - 1; x++)
            {
                var local = y * localWidth + x;
                var source = (minY + y) * imageWidth + minX + x;
                if (!dark[source] && !outside[local])
                {
                    enclosedCount++;
                }
            }
        }

        return enclosedCount >= Math.Max(4, localWidth * localHeight / 100);

        void Enqueue(int x, int y)
        {
            if (x < 0 || y < 0 || x >= localWidth || y >= localHeight)
            {
                return;
            }

            var local = y * localWidth + x;
            var source = (minY + y) * imageWidth + minX + x;
            if (outside[local] || dark[source])
            {
                return;
            }

            outside[local] = true;
            queue.Enqueue(local);
        }
    }

    private static PdfPageRenderOptions BuildRenderOptions(Windows.Data.Pdf.PdfPage page)
    {
        var dimensions = page.Dimensions.MediaBox;
        var width = Math.Max(1d, dimensions.Width);
        var height = Math.Max(1d, dimensions.Height);
        var maxDimension = Math.Max(1000u, OcrEngine.MaxImageDimension);
        var targetWidth = Math.Min(PreferredRenderWidth, maxDimension);
        var targetHeight = (uint)Math.Ceiling(targetWidth * height / width);
        if (targetHeight > maxDimension)
        {
            targetHeight = maxDimension;
            targetWidth = (uint)Math.Ceiling(targetHeight * width / height);
        }

        return new PdfPageRenderOptions
        {
            DestinationWidth = targetWidth,
            DestinationHeight = targetHeight
        };
    }

    private static async Task<FlexibleBenefitOcrPage> BuildOcrPageAsync(
        OcrEngine engine,
        SoftwareBitmap bitmap,
        int pageNumber)
    {
        var result = await engine.RecognizeAsync(bitmap);
        var page = new FlexibleBenefitOcrPage
        {
            PageNumber = pageNumber,
            PixelWidth = bitmap.PixelWidth,
            PixelHeight = bitmap.PixelHeight,
            Text = Clean(result.Text)
        };

        foreach (var sourceLine in result.Lines)
        {
            var words = sourceLine.Words
                .Select(word => new FlexibleBenefitOcrWord
                {
                    Text = Clean(word.Text),
                    Bounds = ToBounds(word.BoundingRect)
                })
                .Where(word => word.Text.Length > 0)
                .ToList();
            if (words.Count == 0)
            {
                continue;
            }

            var line = new FlexibleBenefitOcrLine
            {
                Text = Clean(string.Join(" ", words.Select(word => word.Text))),
                Bounds = Union(words.Select(word => word.Bounds))
            };
            line.Words.AddRange(words);
            page.Lines.Add(line);
        }

        return page;
    }

    private static FlexibleBenefitOcrPage BuildMachineReadablePage(
        PdfPigDocument document,
        int pageNumber,
        int pixelWidth,
        int pixelHeight)
    {
        var sourcePage = document.GetPage(pageNumber);
        var words = sourcePage.GetWords()
            .Where(word => !string.IsNullOrWhiteSpace(word.Text))
            .OrderByDescending(word => word.BoundingBox.Bottom)
            .ThenBy(word => word.BoundingBox.Left)
            .ToList();
        var groupedLines = new List<List<Word>>();
        foreach (var word in words)
        {
            var line = groupedLines.FirstOrDefault(candidate =>
                Math.Abs(candidate[0].BoundingBox.Bottom - word.BoundingBox.Bottom) < 3d);
            if (line is null)
            {
                groupedLines.Add([word]);
            }
            else
            {
                line.Add(word);
            }
        }

        var page = new FlexibleBenefitOcrPage
        {
            PageNumber = pageNumber,
            PixelWidth = pixelWidth,
            PixelHeight = pixelHeight,
            Text = Clean(sourcePage.Text)
        };
        foreach (var sourceLine in groupedLines)
        {
            var lineWords = sourceLine
                .OrderBy(word => word.BoundingBox.Left)
                .Select(word => new FlexibleBenefitOcrWord
                {
                    Text = Clean(word.Text),
                    Bounds = MapPdfBounds(word.BoundingBox, sourcePage.Width, sourcePage.Height, pixelWidth, pixelHeight)
                })
                .Where(word => word.Text.Length > 0)
                .ToList();
            if (lineWords.Count == 0)
            {
                continue;
            }

            var line = new FlexibleBenefitOcrLine
            {
                Text = Clean(string.Join(" ", lineWords.Select(word => word.Text))),
                Bounds = Union(lineWords.Select(word => word.Bounds))
            };
            line.Words.AddRange(lineWords);
            page.Lines.Add(line);
        }

        if (page.Text.Length == 0)
        {
            page = new FlexibleBenefitOcrPage
            {
                PageNumber = pageNumber,
                PixelWidth = pixelWidth,
                PixelHeight = pixelHeight,
                Text = Clean(string.Join(" ", page.Lines.Select(line => line.Text)))
            };
        }

        return page;
    }

    private static IReadOnlyList<FlexibleBenefitDetectedMark> DetectSelectionMarks(
        SoftwareBitmap bitmap,
        IReadOnlyList<FlexibleBenefitOcrLine> lines)
    {
        var width = bitmap.PixelWidth;
        var height = bitmap.PixelHeight;
        var buffer = new Windows.Storage.Streams.Buffer(checked((uint)(width * height * 4)));
        bitmap.CopyToBuffer(buffer);
        var pixels = new byte[buffer.Length];
        using (var reader = DataReader.FromBuffer(buffer))
        {
            reader.ReadBytes(pixels);
        }

        var dark = new bool[width * height];
        var outline = new bool[width * height];
        for (var index = 0; index < dark.Length; index++)
        {
            var offset = index * 4;
            var blue = pixels[offset];
            var green = pixels[offset + 1];
            var red = pixels[offset + 2];
            var luminance = (red + green + blue) / 3;
            dark[index] = luminance < 105;
            outline[index] = luminance < 215;
        }

        var visited = new bool[dark.Length];
        var marks = new List<FlexibleBenefitDetectedMark>();
        var minimumSize = Math.Max(10, (int)Math.Round(width * 0.006));
        var maximumSize = Math.Max(45, (int)Math.Round(width * 0.04));
        var queue = new Queue<int>();

        for (var index = 0; index < dark.Length; index++)
        {
            if (!outline[index] || visited[index])
            {
                continue;
            }

            visited[index] = true;
            queue.Enqueue(index);
            var minX = width;
            var minY = height;
            var maxX = 0;
            var maxY = 0;
            var pixelCount = 0;

            while (queue.Count > 0)
            {
                var current = queue.Dequeue();
                var x = current % width;
                var y = current / width;
                minX = Math.Min(minX, x);
                minY = Math.Min(minY, y);
                maxX = Math.Max(maxX, x);
                maxY = Math.Max(maxY, y);
                pixelCount++;

                Visit(x - 1, y);
                Visit(x + 1, y);
                Visit(x, y - 1);
                Visit(x, y + 1);
            }

            var componentWidth = maxX - minX + 1;
            var componentHeight = maxY - minY + 1;
            var ratio = componentHeight == 0 ? 0d : (double)componentWidth / componentHeight;
            if (componentWidth < minimumSize || componentHeight < minimumSize ||
                componentWidth > maximumSize || componentHeight > maximumSize ||
                ratio < 0.70 || ratio > 1.30 || pixelCount < minimumSize * 2)
            {
                continue;
            }

            var outlineScore = OutlineDarkScore(outline, width, height, minX, minY, maxX, maxY);
            var componentFill = (double)pixelCount / (componentWidth * componentHeight);
            if (outlineScore < 0.24 || componentFill > 0.55)
            {
                continue;
            }

            var fillRatio = InteriorDarkRatio(dark, width, height, minX, minY, maxX, maxY);
            var selected = fillRatio switch
            {
                >= 0.035 => true,
                <= 0.012 => false,
                _ => (bool?)null
            };
            marks.Add(new FlexibleBenefitDetectedMark
            {
                Bounds = new FlexibleBenefitBounds(minX, minY, componentWidth, componentHeight),
                IsSelected = selected,
                Shape = "Selection mark",
                Confidence = selected.HasValue ? 0.94m : 0.55m
            });

            void Visit(int nextX, int nextY)
            {
                if (nextX < 0 || nextY < 0 || nextX >= width || nextY >= height)
                {
                    return;
                }

                var next = nextY * width + nextX;
                if (!outline[next] || visited[next])
                {
                    return;
                }

                visited[next] = true;
                queue.Enqueue(next);
            }
        }

        AddSignatureMarks(dark, width, height, lines, marks);
        return marks;
    }

    private static void AddSignatureMarks(
        IReadOnlyList<bool> dark,
        int width,
        int height,
        IReadOnlyList<FlexibleBenefitOcrLine> lines,
        ICollection<FlexibleBenefitDetectedMark> marks)
    {
        foreach (var line in lines.Where(line =>
                     FlexibleBenefitNormalizationService.NormalizeSearchText(line.Text)
                         .Contains("signatureof", StringComparison.Ordinal)))
        {
            var region = new FlexibleBenefitBounds(
                Math.Max(0d, line.Bounds.X - width * 0.15),
                Math.Max(0d, line.Bounds.Y - height * 0.10),
                Math.Min(width, line.Bounds.Right + width * 0.15) - Math.Max(0d, line.Bounds.X - width * 0.15),
                Math.Max(1d, line.Bounds.Y - height * 0.02 - Math.Max(0d, line.Bounds.Y - height * 0.10)));
            var ratio = DarkRatio(
                dark,
                width,
                height,
                (int)region.X,
                (int)region.Y,
                (int)region.Right,
                (int)region.Bottom);
            var state = ratio switch
            {
                >= 0.010 => true,
                <= 0.002 => false,
                _ => (bool?)null
            };
            marks.Add(new FlexibleBenefitDetectedMark
            {
                Bounds = region,
                IsSelected = state,
                Shape = "Signature ink",
                Confidence = state.HasValue ? 0.90m : 0.55m
            });
        }
    }

    private static double OutlineDarkScore(
        IReadOnlyList<bool> dark,
        int width,
        int height,
        int minX,
        int minY,
        int maxX,
        int maxY)
    {
        var componentWidth = maxX - minX + 1;
        var componentHeight = maxY - minY + 1;
        var thickness = Math.Max(2, Math.Min(componentWidth, componentHeight) / 8);
        var top = DarkRatio(dark, width, height, minX, minY, maxX, Math.Min(maxY, minY + thickness - 1));
        var bottom = DarkRatio(dark, width, height, minX, Math.Max(minY, maxY - thickness + 1), maxX, maxY);
        var left = DarkRatio(dark, width, height, minX, minY, Math.Min(maxX, minX + thickness - 1), maxY);
        var right = DarkRatio(dark, width, height, Math.Max(minX, maxX - thickness + 1), minY, maxX, maxY);
        if (Math.Min(Math.Min(top, bottom), Math.Min(left, right)) < 0.08)
        {
            return 0d;
        }

        return (top + bottom + left + right) / 4d;
    }

    private static double DarkRatio(
        IReadOnlyList<bool> dark,
        int width,
        int height,
        int minX,
        int minY,
        int maxX,
        int maxY)
    {
        minX = Math.Clamp(minX, 0, width - 1);
        maxX = Math.Clamp(maxX, 0, width - 1);
        minY = Math.Clamp(minY, 0, height - 1);
        maxY = Math.Clamp(maxY, 0, height - 1);
        if (minX > maxX || minY > maxY)
        {
            return 0d;
        }

        var total = 0;
        var darkCount = 0;
        for (var y = minY; y <= maxY; y++)
        {
            for (var x = minX; x <= maxX; x++)
            {
                total++;
                if (dark[y * width + x])
                {
                    darkCount++;
                }
            }
        }

        return total == 0 ? 0d : (double)darkCount / total;
    }

    private static double InteriorDarkRatio(
        IReadOnlyList<bool> dark,
        int width,
        int height,
        int minX,
        int minY,
        int maxX,
        int maxY)
    {
        var insetX = Math.Max(2, (maxX - minX + 1) / 5);
        var insetY = Math.Max(2, (maxY - minY + 1) / 5);
        var left = Math.Clamp(minX + insetX, 0, width - 1);
        var right = Math.Clamp(maxX - insetX, 0, width - 1);
        var top = Math.Clamp(minY + insetY, 0, height - 1);
        var bottom = Math.Clamp(maxY - insetY, 0, height - 1);
        if (left > right || top > bottom)
        {
            return 0d;
        }

        var total = 0;
        var darkCount = 0;
        for (var y = top; y <= bottom; y++)
        {
            for (var x = left; x <= right; x++)
            {
                total++;
                if (dark[y * width + x])
                {
                    darkCount++;
                }
            }
        }

        return total == 0 ? 0d : (double)darkCount / total;
    }

    private static bool HasMachineReadableText(PdfPigDocument? document, int pageNumber)
    {
        if (document is null)
        {
            return false;
        }

        try
        {
            var page = document.GetPage(pageNumber);
            return Clean(page.Text).Length >= 30 && page.GetWords().Count() >= 5;
        }
        catch
        {
            return false;
        }
    }

    private static PdfPigDocument? TryOpenPdfPig(string path)
    {
        try
        {
            return PdfPigDocument.Open(path);
        }
        catch
        {
            return null;
        }
    }

    private static FlexibleBenefitBounds MapPdfBounds(
        UglyToad.PdfPig.Core.PdfRectangle bounds,
        double pageWidth,
        double pageHeight,
        int pixelWidth,
        int pixelHeight)
    {
        var x = bounds.Left / pageWidth * pixelWidth;
        var y = (pageHeight - bounds.Top) / pageHeight * pixelHeight;
        var width = bounds.Width / pageWidth * pixelWidth;
        var height = bounds.Height / pageHeight * pixelHeight;
        return new FlexibleBenefitBounds(x, y, width, height);
    }

    private static FlexibleBenefitBounds ToBounds(Windows.Foundation.Rect rect)
    {
        return new FlexibleBenefitBounds(rect.X, rect.Y, rect.Width, rect.Height);
    }

    private static FlexibleBenefitBounds Union(IEnumerable<FlexibleBenefitBounds> bounds)
    {
        var values = bounds.ToList();
        if (values.Count == 0)
        {
            return default;
        }

        var left = values.Min(value => value.X);
        var top = values.Min(value => value.Y);
        var right = values.Max(value => value.Right);
        var bottom = values.Max(value => value.Bottom);
        return new FlexibleBenefitBounds(left, top, right - left, bottom - top);
    }

    private static async Task<string> ComputeHashAsync(string path, CancellationToken cancellationToken)
    {
        await using var stream = File.OpenRead(path);
        var bytes = await SHA256.HashDataAsync(stream, cancellationToken);
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    private static string Clean(string? value)
    {
        return string.Join(" ", (value ?? "")
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
            .Trim();
    }
}
