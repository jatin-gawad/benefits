using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.FlexibleBenefits;

public interface IFlexibleBenefitDocumentParser
{
    string Carrier { get; }
    bool CanParse(FlexibleBenefitOcrDocument document);
    Task<FlexibleBenefitExtractionResult> ParseAsync(
        FlexibleBenefitOcrDocument document,
        CancellationToken cancellationToken = default);
}

public sealed class NationalHealthClaimFlexibleBenefitParser : IFlexibleBenefitDocumentParser
{
    private const string PlanEntityKey = "plan";
    private readonly IFlexibleBenefitDocumentTextExtractor _textExtractor;
    private readonly FlexibleBenefitCategoryCatalogue _categoryCatalogue;
    private readonly FlexibleBenefitNormalizationService _normalizer;
    private readonly FlexibleBenefitLayoutReader _layout;
    private readonly FlexibleBenefitDocumentClassifier _classifier;

    public NationalHealthClaimFlexibleBenefitParser(
        IFlexibleBenefitDocumentTextExtractor? textExtractor = null,
        FlexibleBenefitCategoryCatalogue? categoryCatalogue = null,
        FlexibleBenefitNormalizationService? normalizer = null,
        FlexibleBenefitLayoutReader? layout = null,
        FlexibleBenefitDocumentClassifier? classifier = null)
    {
        _textExtractor = textExtractor ?? new WindowsOcrPdfExtractor();
        _categoryCatalogue = categoryCatalogue ?? new FlexibleBenefitCategoryCatalogue();
        _normalizer = normalizer ?? new FlexibleBenefitNormalizationService();
        _layout = layout ?? new FlexibleBenefitLayoutReader();
        _classifier = classifier ?? new FlexibleBenefitDocumentClassifier();
    }

    public string Carrier => "National HealthClaim";

    public bool CanParse(FlexibleBenefitOcrDocument document)
    {
        var text = FlexibleBenefitNormalizationService.NormalizeSearchText(document.CombinedText);
        return text.Contains("nationalhealthclaim", StringComparison.Ordinal) &&
               (text.Contains("lifehealth", StringComparison.Ordinal) ||
                text.Contains("spendingaccount", StringComparison.Ordinal));
    }

    public async Task<FlexibleBenefitExtractionResult> ParseAsync(
        FlexibleBenefitOcrDocument document,
        CancellationToken cancellationToken = default)
    {
        if (!CanParse(document))
        {
            throw new InvalidDataException("The PDF is not recognized as a National HealthClaim flexible-benefits document.");
        }

        var result = new FlexibleBenefitExtractionResult
        {
            Classification = _classifier.Classify(document),
            OcrDocument = document,
            SourcePath = document.FilePath,
            SourceHash = document.FileHash,
            Confidence = document.Confidence,
            Status = FlexibleBenefitExtractionStatuses.Extracted
        };

        ExtractClientIdentity(result);
        await ExtractPlanAsync(result, cancellationToken);
        ExtractClasses(result);
        ExtractCategories(result);
        ExtractFunding(result);
        await ExtractFeesAsync(result, cancellationToken);
        ExtractAgreement(result);
        ExtractContacts(result);
        ApplyPlanRollups(result);
        result.Status = FlexibleBenefitExtractionStatuses.Normalized;
        return result;
    }

    private void ExtractClientIdentity(FlexibleBenefitExtractionResult result)
    {
        var document = result.OcrDocument!;
        var subscriberLabel = _layout.FindLine(document, "the Subscriber");
        FlexibleBenefitLineMatch? subscriberName = null;
        if (subscriberLabel is not null)
        {
            subscriberName = subscriberLabel.Page.Lines
                .Where(line => line.Bounds.Y < subscriberLabel.Line.Bounds.Y)
                .Where(line => subscriberLabel.Line.Bounds.Y - line.Bounds.Y < subscriberLabel.Page.PixelHeight * 0.08)
                .Where(line => line.Text.Contains("DBA", StringComparison.OrdinalIgnoreCase) || line.Text.Length >= 12)
                .OrderBy(line => subscriberLabel.Line.Bounds.Y - line.Bounds.Y)
                .Select(line => new FlexibleBenefitLineMatch(subscriberLabel.Page, line, "Subscriber", 0.95m))
                .FirstOrDefault();
        }

        var companyLabel = _layout.FindLine(document, "Legal Name of Company");
        var companyValue = companyLabel is null ? null : FindValueRightOfLabel(companyLabel, text => text.Length >= 4);
        var companyPage = companyLabel?.Page ?? document.Pages.FirstOrDefault();
        var completeCompanyName = FirstNonBlank(subscriberName?.Line.Text, companyValue?.Line.Text);
        var legalName = completeCompanyName;
        var dbaName = "";
        var dbaMatch = Regex.Match(completeCompanyName, @"^(?<legal>.+?)\s+DBA\s+(?<dba>.+)$", RegexOptions.IgnoreCase);
        if (dbaMatch.Success)
        {
            legalName = dbaMatch.Groups["legal"].Value.Trim();
            dbaName = dbaMatch.Groups["dba"].Value.Trim();
        }

        var addressLabel = companyPage is null ? null : BestLine(companyPage, "Address");
        var addressValue = addressLabel is null ? null : FindValueRightOfLabel(addressLabel, LooksLikeAddress);
        var provinceLabel = companyPage is null ? null : BestLine(companyPage, "Province");
        var provinceValue = provinceLabel is null ? null : FindValueRightOfLabel(provinceLabel, LooksLikeProvince);
        var cityLine = companyPage is null ? null : BestLine(companyPage, "City");
        var cityValue = cityLine is null ? null : FindValueRightOfLabel(cityLine, text => text.Length >= 3 && text.All(character => char.IsLetter(character) || char.IsWhiteSpace(character)));
        var city = FirstNonBlank(cityValue?.Line.Text, cityLine is null ? "" : ValueAfterLabel(cityLine.Line.Text, "City"));
        var phoneLine = companyPage is null ? null : BestLine(companyPage, "Phone");
        var phoneValue = phoneLine is null
            ? null
            : FindValueRightOfLabel(phoneLine, text => Regex.IsMatch(text, @"\d{3}.*\d{3}.*\d{4}"));
        var phoneRaw = FirstNonBlank(phoneValue?.Line.Text, phoneLine?.Line.Text);
        var phone = Regex.Match(phoneRaw, @"(?:\+?1\s*)?\(?\d{3}\)?[-\s]\d{3}[-\s]\d{4}").Value;
        var postalLine = companyPage is null ? null : BestLine(companyPage, "Postal Code");
        var postalValue = postalLine is null ? null : FindValueRightOfLabel(postalLine, text => text.Length is >= 2 and <= 10);
        var postalRaw = FirstNonBlank(postalValue?.Line.Text, postalLine?.Line.Text);
        var postal = Regex.Match(postalRaw, @"[A-Z]\d[A-Z]\s?\d[A-Z]\d", RegexOptions.IgnoreCase).Value.ToUpperInvariant();

        result.Client = new FlexibleBenefitClientIdentity
        {
            LegalName = _normalizer.NormalizeLegalName(legalName),
            DbaName = _normalizer.NormalizeLegalName(dbaName),
            Address = addressValue?.Line.Text ?? "",
            City = city,
            Province = provinceValue?.Line.Text ?? "",
            PostalCode = postal,
            Phone = phone
        };

        AddSource(result, "Client", "client", "LegalName", completeCompanyName, result.Client.LegalName, subscriberName ?? companyValue, null, 0.96m);
        if (!string.IsNullOrWhiteSpace(result.Client.DbaName))
        {
            AddSource(result, "Client", "client", "DbaName", completeCompanyName, result.Client.DbaName, subscriberName ?? companyValue, null, 0.96m);
        }

        AddSource(result, "Client", "client", "Address", addressValue?.Line.Text, result.Client.Address, addressValue, null, 0.91m);
        AddSource(result, "Client", "client", "City", cityValue?.Line.Text ?? cityLine?.Line.Text, result.Client.City, cityValue ?? cityLine, null, 0.91m);
        AddSource(result, "Client", "client", "Province", provinceValue?.Line.Text, result.Client.Province, provinceValue, null, 0.91m);
        AddSource(result, "Client", "client", "PostalCode", postalRaw, result.Client.PostalCode, postalValue ?? postalLine, null, string.IsNullOrWhiteSpace(postal) ? 0.45m : 0.90m,
            string.IsNullOrWhiteSpace(postal) ? FlexibleBenefitReviewStatuses.NeedsReview : FlexibleBenefitReviewStatuses.Pending);
        AddSource(result, "Client", "client", "Phone", phoneRaw, result.Client.Phone, phoneValue ?? phoneLine, null,
            string.IsNullOrWhiteSpace(phone) ? 0.45m : 0.90m,
            string.IsNullOrWhiteSpace(phone) ? FlexibleBenefitReviewStatuses.NeedsReview : FlexibleBenefitReviewStatuses.Pending);
    }

    private async Task ExtractPlanAsync(FlexibleBenefitExtractionResult result, CancellationToken cancellationToken)
    {
        var document = result.OcrDocument!;
        var combinationLine = _layout.FindLine(document, "combination of Health Spending Account and Lifestyle Spending Account");
        var hcsaLine = _layout.FindLine(document, "comprehensive health spending account coverage");
        var planPage = document.Pages.FirstOrDefault(page => Contains(page.Text, "Default Ratio"));
        var hcsaEnabled = combinationLine is not null || hcsaLine is not null;
        var lsaEnabled = document.Pages.Take(4).Any(page => Contains(page.Text, "Lifestyle Spending Account"));
        var plan = new FlexibleBenefitPlan
        {
            PlanName = "LifeHealth+",
            HcsaEnabled = hcsaEnabled ? true : null,
            LsaEnabled = lsaEnabled ? true : null,
            CombinedAccountEnabled = hcsaEnabled && lsaEnabled ? true : null,
            WellnessAccountEnabled = false,
            CurrencyCode = "CAD",
            TaxTreatment = "HCSA PHSP and taxable LSA",
            ReviewStatus = FlexibleBenefitReviewStatuses.Pending,
            ExtractionStatus = FlexibleBenefitExtractionStatuses.Normalized
        };
        plan.PlanType = _normalizer.NormalizePlanType(plan.HcsaEnabled, plan.LsaEnabled, document.CombinedText);

        AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "PlanName", "LifeHealth+", plan.PlanName, _layout.FindLine(document, "LifeHealth+ Enrollment Form"), null, 0.96m);
        AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "HcsaEnabled", hcsaLine?.Line.Text ?? combinationLine?.Line.Text, BoolText(plan.HcsaEnabled), hcsaLine ?? combinationLine, null, 0.93m);
        AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "LsaEnabled", combinationLine?.Line.Text, BoolText(plan.LsaEnabled), combinationLine, null, 0.93m);
        AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "PlanType", combinationLine?.Line.Text, plan.PlanType.ToString(), combinationLine, null, 0.91m);

        if (planPage is not null)
        {
            var ratio = _layout.FindOrdinalSelection(planPage, "Default Ratio", ["100% HSA", "100% LSA", "50% HSA/LSA"]);
            if (!string.IsNullOrWhiteSpace(ratio?.SelectedValue))
            {
                var normalizedRatio = _normalizer.NormalizeRatio(ratio.SelectedValue);
                plan.DefaultAllocationType = normalizedRatio.Type;
                plan.DefaultHcsaPercentage = normalizedRatio.HcsaPercentage;
                plan.DefaultLsaPercentage = normalizedRatio.LsaPercentage;
                AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "DefaultAllocationType", ratio.SelectedValue, plan.DefaultAllocationType.ToString(), ratio.Anchor, ratio.SelectedMark, ratio.Confidence);
                AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "DefaultHcsaPercentage", ratio.SelectedValue, DecimalText(plan.DefaultHcsaPercentage), ratio.Anchor, ratio.SelectedMark, ratio.Confidence);
                AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "DefaultLsaPercentage", ratio.SelectedValue, DecimalText(plan.DefaultLsaPercentage), ratio.Anchor, ratio.SelectedMark, ratio.Confidence);
            }

            var effectiveDateLine = planPage.Lines
                .Select(line => new { Line = line, Date = _normalizer.ParseDate(line.Text) })
                .FirstOrDefault(item => item.Date.HasValue && item.Date.Value.Year >= 2000);
            plan.EffectiveDate = effectiveDateLine?.Date;
            var effectiveMatch = effectiveDateLine is null ? null : new FlexibleBenefitLineMatch(planPage, effectiveDateLine.Line, "Plan Effective Date", 0.95m);
            AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "EffectiveDate", effectiveDateLine?.Line.Text, DateText(plan.EffectiveDate), effectiveMatch, null, plan.EffectiveDate.HasValue ? 0.96m : 0.45m);

            var calendarSelection = _layout.FindSelection(planPage, "January to December");
            plan.BenefitYearType = calendarSelection?.Mark.IsSelected == true
                ? FlexibleBenefitYearType.CalendarYear
                : FlexibleBenefitYearType.Unknown;
            if (plan.BenefitYearType == FlexibleBenefitYearType.CalendarYear)
            {
                plan.BenefitYearStartMonth = 1;
                plan.BenefitYearStartDay = 1;
            }

            AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "BenefitYearType", calendarSelection?.LineMatch.Line.Text, plan.BenefitYearType.ToString(), calendarSelection?.LineMatch, calendarSelection?.Mark, calendarSelection?.Mark.Confidence ?? 0.45m);

            var runoff = _layout.FindOrdinalSelection(planPage, "Run-off", ["30", "60", "90"]);
            plan.RunoffDays = runoff?.SelectedValue is null ? null : _normalizer.ParseDays(runoff.SelectedValue);
            AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "RunoffDays", runoff?.SelectedValue, IntegerText(plan.RunoffDays), runoff?.Anchor, runoff?.SelectedMark, runoff?.Confidence ?? 0.45m);

            var carryChoices = new[]
            {
                (Text: "Use Credit Carry Forward", Type: FlexibleBenefitCarryForwardType.Credit),
                (Text: "Use Expense Carry Forward", Type: FlexibleBenefitCarryForwardType.Expense),
                (Text: "Use No Carry Forwards", Type: FlexibleBenefitCarryForwardType.None)
            };
            var selectedCarry = carryChoices
                .Select(choice =>
                {
                    var line = BestLine(planPage, choice.Text);
                    var mark = line is null ? null : FindMarkForLine(planPage, line.Line);
                    return (choice, line, mark);
                })
                .FirstOrDefault(item => item.mark?.IsSelected == true);
            if (selectedCarry.line is not null && selectedCarry.mark is not null)
            {
                plan.CarryForwardType = selectedCarry.choice.Type;
                plan.CreditCarryForwardEnabled = selectedCarry.choice.Type == FlexibleBenefitCarryForwardType.Credit;
                plan.ExpenseCarryForwardEnabled = selectedCarry.choice.Type == FlexibleBenefitCarryForwardType.Expense;
                AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "CarryForwardType", selectedCarry.line.Line.Text, plan.CarryForwardType.ToString(), selectedCarry.line, selectedCarry.mark, selectedCarry.mark.Confidence);
            }

            var childLabel = BestLine(planPage, "Child Dependent Cut-Off Age");
            var studentLabel = BestLine(planPage, "Student Dependent Cut-Off Age");
            var childValue = await ExtractIntegerRightOfLabelAsync(document, childLabel, cancellationToken);
            var studentValue = await ExtractIntegerRightOfLabelAsync(document, studentLabel, cancellationToken);
            plan.ChildDependentCutoffAge = childValue.Value;
            plan.StudentDependentCutoffAge = studentValue.Value;
            AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "ChildDependentCutoffAge", childValue.RawValue, IntegerText(plan.ChildDependentCutoffAge), childValue.Match ?? childLabel, null, childValue.Confidence);
            AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "StudentDependentCutoffAge", studentValue.RawValue, IntegerText(plan.StudentDependentCutoffAge), studentValue.Match ?? studentLabel, null, studentValue.Confidence);
        }

        var payPage = document.Pages.FirstOrDefault(page => Contains(page.Text, "PAY AS SUBMITTED"));
        var paySelection = payPage is null ? null : _layout.FindSelection(payPage, "Pay all receipts as submitted");
        plan.PayAsSubmitted = paySelection?.Mark.IsSelected;
        AddSource(result, "FlexibleBenefitPlan", PlanEntityKey, "PayAsSubmitted", paySelection?.LineMatch.Line.Text, BoolText(plan.PayAsSubmitted), paySelection?.LineMatch, paySelection?.Mark, paySelection?.Mark.Confidence ?? 0.45m,
            plan.PayAsSubmitted.HasValue ? FlexibleBenefitReviewStatuses.Pending : FlexibleBenefitReviewStatuses.NeedsReview);

        result.Plan = plan;
    }

    private void ExtractClasses(FlexibleBenefitExtractionResult result)
    {
        var document = result.OcrDocument!;
        var page = document.Pages.FirstOrDefault(item => Contains(item.Text, "JOB CLASSIFICATION"));
        if (page is null)
        {
            return;
        }

        var singleHeader = FindHeader(page, "SINGLE");
        var coupleHeader = FindHeader(page, "COUPLE");
        var familyHeader = FindHeader(page, "FAMILY");
        var waitingHeader = FindHeader(page, "WAITING");
        var defaultRatio = _layout.FindLines(page, "Default Ratio").FirstOrDefault();
        if (singleHeader is null || coupleHeader is null || familyHeader is null || waitingHeader is null || defaultRatio is null)
        {
            return;
        }

        var tableTop = new[] { singleHeader, coupleHeader, familyHeader, waitingHeader }.Max(item => item.Line.Bounds.Bottom);
        var numericRows = page.Lines
            .Where(line => line.Bounds.Y > tableTop)
            .Where(line => line.Bounds.Y < defaultRatio.Line.Bounds.Y - page.PixelHeight * 0.03)
            .Where(line => _normalizer.ParseMoney(line.Text).HasValue)
            .GroupBy(line => (int)Math.Round(line.Bounds.Y / Math.Max(20d, page.PixelHeight * 0.025)))
            .Select(group => group.ToList())
            .Where(group => group.Count >= 3)
            .ToList();

        foreach (var row in numericRows)
        {
            var rowY = row.Average(line => line.Bounds.Y);
            var classNameLine = page.Lines
                .Where(line => line.Bounds.X < singleHeader.Line.Bounds.X * 0.85)
                .Where(line => Math.Abs(line.Bounds.Y - rowY) < page.PixelHeight * 0.035)
                .Where(line => !Contains(line.Text, "YES"))
                .OrderBy(line => Math.Abs(line.Bounds.Y - rowY))
                .FirstOrDefault();
            if (classNameLine is null || string.IsNullOrWhiteSpace(classNameLine.Text))
            {
                continue;
            }

            var singleLine = ClosestColumnValue(row, singleHeader.Line.Bounds.X);
            var coupleLine = ClosestColumnValue(row, coupleHeader.Line.Bounds.X);
            var familyLine = ClosestColumnValue(row, familyHeader.Line.Bounds.X);
            var waitingLine = page.Lines
                .Where(line => Math.Abs(line.Bounds.Y - rowY) < page.PixelHeight * 0.035)
                .Where(line => line.Bounds.X >= waitingHeader.Line.Bounds.X - page.PixelWidth * 0.03)
                .Where(line => _normalizer.ParseDays(line.Text).HasValue)
                .OrderBy(line => Math.Abs(line.Bounds.X - waitingHeader.Line.Bounds.X))
                .FirstOrDefault();
            var yesLine = page.Lines
                .Where(line => Contains(line.Text, "YES"))
                .Where(line => Math.Abs(line.Bounds.Y - rowY) < page.PixelHeight * 0.04)
                .OrderBy(line => Math.Abs(line.Bounds.Y - rowY))
                .FirstOrDefault();
            var prorationMark = yesLine is null ? null : FindMarkForLine(page, yesLine);
            var item = new FlexibleBenefitClass
            {
                ClassName = classNameLine.Text.Trim(),
                AnnualSingleAmount = _normalizer.ParseMoney(singleLine?.Text),
                AnnualCoupleAmount = _normalizer.ParseMoney(coupleLine?.Text),
                AnnualFamilyAmount = _normalizer.ParseMoney(familyLine?.Text),
                WaitingPeriodDays = _normalizer.ParseDays(waitingLine?.Text),
                NewHireProrationEnabled = prorationMark?.IsSelected,
                HcsaPercentage = result.Plan.DefaultHcsaPercentage,
                LsaPercentage = result.Plan.DefaultLsaPercentage,
                CurrencyCode = "CAD",
                EffectiveDate = result.Plan.EffectiveDate,
                SourcePageNumber = page.PageNumber,
                ExtractionConfidence = 0.92m,
                ReviewStatus = prorationMark?.IsSelected.HasValue == true
                    ? FlexibleBenefitReviewStatuses.Pending
                    : FlexibleBenefitReviewStatuses.NeedsReview
            };
            result.Classes.Add(item);

            AddSource(result, "FlexibleBenefitClass", item.SourceKey, "ClassName", classNameLine.Text, item.ClassName, Match(page, classNameLine, "Class Name"), null, 0.94m);
            AddSource(result, "FlexibleBenefitClass", item.SourceKey, "AnnualSingleAmount", singleLine?.Text, DecimalText(item.AnnualSingleAmount), Match(page, singleLine, "Annual Single"), null, 0.95m);
            AddSource(result, "FlexibleBenefitClass", item.SourceKey, "AnnualCoupleAmount", coupleLine?.Text, DecimalText(item.AnnualCoupleAmount), Match(page, coupleLine, "Annual Couple"), null, 0.95m);
            AddSource(result, "FlexibleBenefitClass", item.SourceKey, "AnnualFamilyAmount", familyLine?.Text, DecimalText(item.AnnualFamilyAmount), Match(page, familyLine, "Annual Family"), null, 0.95m);
            AddSource(result, "FlexibleBenefitClass", item.SourceKey, "WaitingPeriodDays", waitingLine?.Text, IntegerText(item.WaitingPeriodDays), Match(page, waitingLine, "Waiting Period"), null, 0.94m);
            AddSource(result, "FlexibleBenefitClass", item.SourceKey, "NewHireProrationEnabled", yesLine?.Text, BoolText(item.NewHireProrationEnabled), Match(page, yesLine, "New Hire Proration"), prorationMark, prorationMark?.Confidence ?? 0.45m,
                item.NewHireProrationEnabled.HasValue ? FlexibleBenefitReviewStatuses.Pending : FlexibleBenefitReviewStatuses.NeedsReview);

            if (item.WaitingPeriodDays.HasValue)
            {
                var rule = new FlexibleBenefitEligibilityRule
                {
                    ClassSourceKey = item.SourceKey,
                    RuleType = "WaitingPeriod",
                    Operator = "Equals",
                    RuleValue = item.WaitingPeriodDays.Value.ToString(CultureInfo.InvariantCulture),
                    RuleUnit = "Days",
                    Description = $"{item.WaitingPeriodDays.Value} day waiting period",
                    EffectiveDate = item.EffectiveDate,
                    SourcePageNumber = item.SourcePageNumber,
                    ExtractionConfidence = item.ExtractionConfidence,
                    ReviewStatus = item.ReviewStatus
                };
                result.EligibilityRules.Add(rule);
                AddSource(result, "FlexibleBenefitEligibilityRule", rule.SourceKey, "RuleValue", waitingLine?.Text, rule.RuleValue, Match(page, waitingLine, "Waiting Period"), null, 0.94m);
            }
        }

        if (result.Classes.Count > 0)
        {
            var electionLine = _layout.FindLine(document, "do not participate in the flex selection process");
            var rule = new FlexibleBenefitEligibilityRule
            {
                RuleType = "AccountElection",
                Operator = "Allows",
                RuleValue = "EmployeeElectionWithEmployerDefault",
                Description = "Employees may participate in the flex selection process; non-participants receive the employer default allocation.",
                EffectiveDate = result.Plan.EffectiveDate,
                SourcePageNumber = electionLine?.Page.PageNumber,
                ExtractionConfidence = 0.90m,
                ReviewStatus = FlexibleBenefitReviewStatuses.Pending
            };
            result.EligibilityRules.Add(rule);
            AddSource(result, "FlexibleBenefitEligibilityRule", rule.SourceKey, "RuleValue", electionLine?.Line.Text, rule.RuleValue, electionLine, null, 0.90m);
        }
    }

    private void ExtractCategories(FlexibleBenefitExtractionResult result)
    {
        var pages = result.OcrDocument!.Pages.Take(4).ToList();
        var definitions = _categoryCatalogue.Load();
        var headings = definitions.Select(definition => new
        {
            Definition = definition,
            Match = pages
                .SelectMany(page => _layout.FindLines(page, definition.DocumentWording))
                .Where(match => match.MatchScore >= 0.70m)
                .OrderByDescending(match => match.MatchScore)
                .ThenBy(match => match.Page.PageNumber)
                .ThenBy(match => match.Line.Bounds.Y)
                .FirstOrDefault()
                ?? pages
                    .SelectMany(page => _layout.FindLines(page, definition.Aliases.ToArray()))
                    .OrderByDescending(match => match.MatchScore)
                    .ThenBy(match => match.Page.PageNumber)
                    .ThenBy(match => match.Line.Bounds.Y)
                    .FirstOrDefault()
        }).ToList();

        foreach (var heading in headings)
        {
            var definition = heading.Definition;
            var nextHeadingY = heading.Match is null
                ? (double?)null
                : headings
                    .Where(item => item.Match?.Page.PageNumber == heading.Match.Page.PageNumber)
                    .Where(item => item.Match!.Line.Bounds.Y > heading.Match.Line.Bounds.Y)
                    .Select(item => (double?)item.Match!.Line.Bounds.Y)
                    .OrderBy(value => value)
                    .FirstOrDefault();
            var selection = heading.Match is null
                ? null
                : FindCategorySelection(heading.Match, nextHeadingY);

            var item = new FlexibleBenefitCategory
            {
                CategoryCode = definition.CategoryCode,
                CategoryGroup = definition.CategoryGroup,
                CategoryName = definition.CategoryName,
                Description = definition.DocumentWording,
                Enabled = selection?.Mark.IsSelected,
                Taxable = definition.Taxable,
                AccountType = definition.AccountType,
                RequiresReceipt = definition.RequiresReceipt,
                RequiresProviderDetails = definition.RequiresProviderDetails,
                Exclusions = definition.Exclusions,
                SourcePageNumber = selection?.LineMatch.Page.PageNumber,
                ExtractionConfidence = selection?.Mark.Confidence ?? 0.40m,
                ReviewStatus = selection?.Mark.IsSelected.HasValue == true
                    ? FlexibleBenefitReviewStatuses.Pending
                    : FlexibleBenefitReviewStatuses.NeedsReview
            };
            result.Categories.Add(item);
            AddSource(result, "FlexibleBenefitCategory", item.SourceKey, "Enabled", selection?.LineMatch.Line.Text ?? definition.DocumentWording, BoolText(item.Enabled), selection?.LineMatch, selection?.Mark, item.ExtractionConfidence ?? 0.40m, item.ReviewStatus);
        }

        var payCategory = result.Categories.FirstOrDefault(category =>
            category.CategoryCode.Equals("PAY_ALL_RECEIPTS_AS_SUBMITTED", StringComparison.OrdinalIgnoreCase));
        if (payCategory is not null)
        {
            result.Plan.PayAsSubmitted = payCategory.Enabled;
        }
    }

    private static FlexibleBenefitSelectionMatch? FindCategorySelection(
        FlexibleBenefitLineMatch heading,
        double? nextHeadingY)
    {
        var page = heading.Page;
        var minimumSize = page.PixelWidth * 0.016;
        var maximumSize = page.PixelWidth * 0.035;
        var startY = heading.Line.Bounds.Y - page.PixelHeight * 0.015;
        var endY = nextHeadingY.HasValue
            ? nextHeadingY.Value - page.PixelHeight * 0.01
            : page.PixelHeight;
        var headingCenterY = heading.Line.Bounds.Y + heading.Line.Bounds.Height / 2d;
        var mark = page.Marks
            .Where(item => item.Shape == "Selection mark")
            .Where(item => item.Bounds.Width >= minimumSize && item.Bounds.Height >= minimumSize)
            .Where(item => item.Bounds.Width <= maximumSize && item.Bounds.Height <= maximumSize)
            .Where(item => item.Bounds.X < heading.Line.Bounds.X + page.PixelWidth * 0.04)
            .Where(item => item.Bounds.Y + item.Bounds.Height / 2d >= startY)
            .Where(item => item.Bounds.Y + item.Bounds.Height / 2d < endY)
            .OrderBy(item => Math.Abs(item.Bounds.Y + item.Bounds.Height / 2d - headingCenterY))
            .FirstOrDefault();
        return mark is null
            ? null
            : new FlexibleBenefitSelectionMatch(
                heading,
                mark,
                Math.Abs(mark.Bounds.Y + mark.Bounds.Height / 2d - headingCenterY));
    }

    private void ExtractFunding(FlexibleBenefitExtractionResult result)
    {
        var document = result.OcrDocument!;
        var padPage = document.Pages.FirstOrDefault(page => Contains(page.Text, "pre-authorized debit"));
        var padSelection = padPage is null
            ? null
            : _layout.FindSelection(padPage, "please contact me about setting up pre-authorized debit");
        var onlineLine = _layout.FindLine(document, "online bank payments or mailing cheques");
        var clearLine = _layout.FindLine(document, "until such funds have actually cleared");
        var interestLine = _layout.FindLine(document, "does not pay interest on monies held");
        var payAsYouGoLine = _layout.FindLine(document, "Pay-as-you-go funding is also available");
        var funding = new FlexibleBenefitFundingConfiguration
        {
            FundingMethod = padSelection?.Mark.IsSelected == true
                ? FlexibleBenefitFundingMethod.Pad
                : FlexibleBenefitFundingMethod.Unknown,
            PreFundingRequired = null,
            PadSelected = padSelection?.Mark.IsSelected,
            ContinuousFloatRequired = padSelection?.Mark.IsSelected == true ? true : null,
            OnlineBankPaymentAllowed = onlineLine is not null ? true : null,
            ChequeAllowed = onlineLine is not null ? true : null,
            PaymentTimingDescription = FirstNonBlank(payAsYouGoLine?.Line.Text, clearLine?.Line.Text),
            ClaimsPaidOnlyAfterFundsClear = clearLine is not null ? true : null,
            InterestPaidOnFunds = interestLine is not null ? false : null,
            CurrencyCode = "CAD",
            EffectiveDate = result.Plan.EffectiveDate,
            SourcePageNumber = padSelection?.LineMatch.Page.PageNumber ?? onlineLine?.Page.PageNumber,
            ExtractionConfidence = padSelection?.Mark.Confidence ?? 0.78m,
            ReviewStatus = padSelection?.Mark.IsSelected.HasValue == true
                ? FlexibleBenefitReviewStatuses.Pending
                : FlexibleBenefitReviewStatuses.NeedsReview
        };
        result.FundingConfigurations.Add(funding);
        AddSource(result, "FlexibleBenefitFundingConfiguration", funding.SourceKey, "PadSelected", padSelection?.LineMatch.Line.Text, BoolText(funding.PadSelected), padSelection?.LineMatch, padSelection?.Mark, padSelection?.Mark.Confidence ?? 0.45m, funding.ReviewStatus);
        AddSource(result, "FlexibleBenefitFundingConfiguration", funding.SourceKey, "OnlineBankPaymentAllowed", onlineLine?.Line.Text, BoolText(funding.OnlineBankPaymentAllowed), onlineLine, null, 0.88m);
        AddSource(result, "FlexibleBenefitFundingConfiguration", funding.SourceKey, "ChequeAllowed", onlineLine?.Line.Text, BoolText(funding.ChequeAllowed), onlineLine, null, 0.88m);
        AddSource(result, "FlexibleBenefitFundingConfiguration", funding.SourceKey, "ClaimsPaidOnlyAfterFundsClear", clearLine?.Line.Text, BoolText(funding.ClaimsPaidOnlyAfterFundsClear), clearLine, null, 0.90m);
        AddSource(result, "FlexibleBenefitFundingConfiguration", funding.SourceKey, "InterestPaidOnFunds", interestLine?.Line.Text, BoolText(funding.InterestPaidOnFunds), interestLine, null, 0.92m);
    }

    private async Task ExtractFeesAsync(
        FlexibleBenefitExtractionResult result,
        CancellationToken cancellationToken)
    {
        var document = result.OcrDocument!;
        var setupLine = _layout.FindLine(document, "Set-up fee");
        FlexibleBenefitOcrLine? setupValueLine = setupLine is null
            ? null
            : setupLine.Page.Lines
                .Where(line => Math.Abs(line.Bounds.Y - setupLine.Line.Bounds.Y) < setupLine.Page.PixelHeight * 0.04)
                .Where(line => line.Bounds.X > setupLine.Line.Bounds.Right)
                .Where(line => line.Bounds.X < setupLine.Line.Bounds.Right + setupLine.Page.PixelWidth * 0.30)
                .Where(line => _normalizer.ParseMoney(line.Text).HasValue)
                .OrderBy(line => line.Bounds.X)
                .FirstOrDefault();
        if (setupLine is not null && setupValueLine is null)
        {
            var region = new FlexibleBenefitBounds(
                Math.Max(0d, setupLine.Line.Bounds.X + setupLine.Line.Bounds.Width * 0.45),
                Math.Max(0d, setupLine.Line.Bounds.Y - setupLine.Page.PixelHeight * 0.012),
                setupLine.Page.PixelWidth * 0.16,
                Math.Max(setupLine.Line.Bounds.Height * 3d, setupLine.Page.PixelHeight * 0.05));
            setupValueLine = (await _textExtractor.ExtractRegionAsync(
                    document.FilePath,
                    setupLine.Page.PageNumber,
                    region,
                    setupLine.Page.PixelWidth,
                    setupLine.Page.PixelHeight,
                    cancellationToken))
                .FirstOrDefault(line => _normalizer.ParseMoney(line.Text).HasValue);
        }

        var setupValue = _normalizer.ParseMoney(setupValueLine?.Text);
        var setupFee = new FlexibleBenefitFee
        {
            FeeType = "SetupFee",
            FeeCalculationType = "FlatAmount",
            FlatAmount = setupValue,
            Taxable = null,
            Description = "Set-up fee",
            EffectiveDate = result.Plan.EffectiveDate,
            SourcePageNumber = setupLine?.Page.PageNumber,
            ExtractionConfidence = setupValue.HasValue ? 0.96m : 0.45m,
            ReviewStatus = setupValue.HasValue ? FlexibleBenefitReviewStatuses.Pending : FlexibleBenefitReviewStatuses.NeedsReview
        };
        result.Fees.Add(setupFee);
        AddSource(result, "FlexibleBenefitFee", setupFee.SourceKey, "FlatAmount", setupValueLine?.Text, DecimalText(setupFee.FlatAmount), setupValueLine is null || setupLine is null ? setupLine : Match(setupLine.Page, setupValueLine, "Set-up fee"), null, setupFee.ExtractionConfidence ?? 0.45m, setupFee.ReviewStatus);

        var adminLabel = _layout.FindLine(document, "Agreed Upon Administration Fee");
        var adminValue = adminLabel is null
            ? null
            : adminLabel.Page.Lines
                .Where(line => line.Bounds.Y < adminLabel.Line.Bounds.Y)
                .Where(line => adminLabel.Line.Bounds.Y - line.Bounds.Y < adminLabel.Page.PixelHeight * 0.08)
                .Select(line => new { Line = line, Value = _normalizer.ParsePercentage(line.Text) })
                .FirstOrDefault(item => item.Value.HasValue);
        var gstLine = _layout.FindLine(document, "GST is charged on the administration fee");
        var noticeLine = document.Pages
            .SelectMany(page => page.Lines.Select(line => new { Page = page, Line = line }))
            .Where(item => Contains(item.Line.Text, "Administrative Fee"))
            .Where(item => Contains(item.Line.Text, "written notice"))
            .Select(item => Match(item.Page, item.Line, "Administrative Fee notice"))
            .FirstOrDefault();
        var noticeDays = noticeLine is null ? null : _normalizer.ParseDays(noticeLine.Line.Text);
        var adminFee = new FlexibleBenefitFee
        {
            FeeType = "AdministrationFee",
            FeeCalculationType = "Percentage",
            Percentage = adminValue?.Value,
            Taxable = gstLine is not null ? true : null,
            TaxType = gstLine is not null ? "GST" : "",
            Description = "Percentage of employee benefits paid out under the LifeHealth+ Account",
            EffectiveDate = result.Plan.EffectiveDate,
            NoticePeriodDays = noticeDays,
            SourcePageNumber = adminLabel?.Page.PageNumber,
            ExtractionConfidence = adminValue?.Value.HasValue == true ? 0.97m : 0.45m,
            ReviewStatus = adminValue?.Value.HasValue == true ? FlexibleBenefitReviewStatuses.Pending : FlexibleBenefitReviewStatuses.NeedsReview
        };
        result.Fees.Add(adminFee);
        AddSource(result, "FlexibleBenefitFee", adminFee.SourceKey, "Percentage", adminValue?.Line.Text, DecimalText(adminFee.Percentage), adminValue is null ? adminLabel : Match(adminLabel!.Page, adminValue.Line, "Administration fee"), null, adminFee.ExtractionConfidence ?? 0.45m, adminFee.ReviewStatus);
        AddSource(result, "FlexibleBenefitFee", adminFee.SourceKey, "TaxType", gstLine?.Line.Text, adminFee.TaxType, gstLine, null, 0.96m);
        AddSource(result, "FlexibleBenefitFee", adminFee.SourceKey, "NoticePeriodDays", noticeLine?.Line.Text, IntegerText(adminFee.NoticePeriodDays), noticeLine, null, noticeDays.HasValue ? 0.94m : 0.45m);
    }

    private void ExtractAgreement(FlexibleBenefitExtractionResult result)
    {
        var document = result.OcrDocument!;
        var agreementLine = _layout.FindLine(document, "Administrative Services Indemnity Agreement");
        var datePage = agreementLine?.Page;
        var monthLine = datePage?.Lines.FirstOrDefault(line => Contains(line.Text, "February") || Contains(line.Text, "January") || Contains(line.Text, "March"));
        var yearLine = datePage?.Lines.FirstOrDefault(line => Regex.IsMatch(line.Text, @"20\s*\d{2}"));
        var dayLine = datePage?.Lines.FirstOrDefault(line => Regex.IsMatch(line.Text, @"the\s+\d{1,2}", RegexOptions.IgnoreCase));
        var agreementDateRaw = string.Join(" ", new[] { monthLine?.Text, dayLine?.Text, yearLine?.Text }.Where(value => !string.IsNullOrWhiteSpace(value)));
        var agreementDate = ParseAgreementDate(agreementDateRaw);
        var signedPage = document.Pages.FirstOrDefault(page => Contains(page.Text, "Signed in (city)"));
        var electronicLine = signedPage?.Lines.FirstOrDefault(line => Contains(line.Text, "E-SIGNED by"));
        var electronicDateLine = signedPage?.Lines.FirstOrDefault(line => Regex.IsMatch(line.Text, @"20\d{2}[-/]\d{2}[-/]\d{2}"));
        var executionDateOnly = _normalizer.ParseDate(electronicDateLine?.Text);
        var subscriberSignatureLine = signedPage is null ? null : _layout.FindLines(signedPage, "Signature of Subscriber").FirstOrDefault();
        var officerSignatureLine = signedPage is null ? null : _layout.FindLines(signedPage, "Signature of Officer").FirstOrDefault();
        var subscriberInk = FindSignatureMark(signedPage, subscriberSignatureLine);
        var officerInk = FindSignatureMark(signedPage, officerSignatureLine);
        var subscriberSigned = electronicLine is not null && executionDateOnly.HasValue ? true : subscriberInk?.IsSelected;
        var officerSigned = officerInk?.IsSelected;
        var cityLine = signedPage?.Lines.FirstOrDefault(line => Contains(line.Text, "Signed in (city)"));
        var city = cityLine is null ? "" : ValueAfterLabel(cityLine.Text, "Signed in (city)");
        var provinceLine = signedPage?.Lines
            .Where(line => line.Bounds.Y >= (cityLine?.Bounds.Y ?? 0) - 20)
            .Where(line => line.Bounds.Y <= (cityLine?.Bounds.Bottom ?? 0) + 40)
            .FirstOrDefault(line => LooksLikeProvince(line.Text));
        var adminSignerLine = FindExplicitSignerLine(signedPage, "Administrator signer name");
        var subscriberSignerLine = FindExplicitSignerLine(signedPage, "Subscriber signer name");
        var adminTitleLine = signedPage?.Lines.FirstOrDefault(line => line.Text.Trim().Equals("Director", StringComparison.OrdinalIgnoreCase));
        var subscriberTitleLine = signedPage?.Lines.FirstOrDefault(line => line.Text.Trim().Equals("Owner", StringComparison.OrdinalIgnoreCase));
        var governingLine = _layout.FindLine(document, "laws of the Province of Alberta");
        var versionLine = _layout.FindLine(document, "LifeHealth+ Indemnity - 1220", "LifeHealth+ Indemnity - 0818");
        var agreement = new FlexibleBenefitAgreement
        {
            AgreementType = "AdministrativeServicesIndemnityAgreement",
            AgreementVersion = versionLine?.Line.Text ?? "",
            AgreementDate = agreementDate,
            ExecutionDate = executionDateOnly?.ToDateTime(TimeOnly.MinValue),
            GoverningProvince = governingLine is null ? "" : "Alberta",
            SubscriberLegalName = result.Client.LegalName + (string.IsNullOrWhiteSpace(result.Client.DbaName) ? "" : $" DBA {result.Client.DbaName}"),
            AdministratorLegalName = "National HealthClaim Corp.",
            SubscriberSignerName = ValueAfterLabel(subscriberSignerLine?.Text ?? "", "Subscriber signer name"),
            SubscriberSignerTitle = subscriberTitleLine?.Text ?? "",
            AdministratorSignerName = ValueAfterLabel(adminSignerLine?.Text ?? "", "Administrator signer name"),
            AdministratorSignerTitle = adminTitleLine?.Text ?? "",
            SubscriberSignaturePresent = subscriberSigned,
            AdministratorSignaturePresent = officerSigned,
            ElectronicSignature = electronicLine is not null ? true : null,
            SignedCity = city,
            SignedProvince = provinceLine?.Text ?? "",
            IndemnityAgreementSigned = subscriberSigned == true && officerSigned == true ? true : null,
            FundingAgreementSigned = null,
            AgreementStatus = subscriberSigned == true && officerSigned == true ? "Signed" : "Needs Review",
            SourcePageNumber = signedPage?.PageNumber,
            ReviewStatus = subscriberSigned == true && officerSigned == true &&
                subscriberSignerLine is not null && adminSignerLine is not null
                ? FlexibleBenefitReviewStatuses.Pending
                : FlexibleBenefitReviewStatuses.NeedsReview
        };
        result.Agreements.Add(agreement);
        AddSource(result, "FlexibleBenefitAgreement", agreement.SourceKey, "AgreementDate", agreementDateRaw, DateText(agreement.AgreementDate), agreementLine, null, agreement.AgreementDate.HasValue ? 0.92m : 0.45m);
        AddSource(result, "FlexibleBenefitAgreement", agreement.SourceKey, "ExecutionDate", electronicDateLine?.Text, agreement.ExecutionDate?.ToString("O", CultureInfo.InvariantCulture), electronicDateLine is null || signedPage is null ? null : Match(signedPage, electronicDateLine, "Execution date"), null, agreement.ExecutionDate.HasValue ? 0.94m : 0.45m);
        AddSource(result, "FlexibleBenefitAgreement", agreement.SourceKey, "GoverningProvince", governingLine?.Line.Text, agreement.GoverningProvince, governingLine, null, 0.96m);
        AddSource(result, "FlexibleBenefitAgreement", agreement.SourceKey, "SubscriberSignerName", subscriberSignerLine?.Text, agreement.SubscriberSignerName, subscriberSignerLine is null || signedPage is null ? null : Match(signedPage, subscriberSignerLine, "Subscriber signer"), null, subscriberSignerLine is null ? 0m : 0.95m, subscriberSignerLine is null ? FlexibleBenefitReviewStatuses.NeedsReview : FlexibleBenefitReviewStatuses.Pending);
        AddSource(result, "FlexibleBenefitAgreement", agreement.SourceKey, "AdministratorSignerName", adminSignerLine?.Text, agreement.AdministratorSignerName, adminSignerLine is null || signedPage is null ? null : Match(signedPage, adminSignerLine, "Administrator signer"), null, adminSignerLine is null ? 0m : 0.95m, adminSignerLine is null ? FlexibleBenefitReviewStatuses.NeedsReview : FlexibleBenefitReviewStatuses.Pending);
        AddSource(result, "FlexibleBenefitAgreement", agreement.SourceKey, "SubscriberSignaturePresent", electronicLine?.Text, BoolText(agreement.SubscriberSignaturePresent), subscriberSignatureLine, subscriberInk, agreement.SubscriberSignaturePresent == true ? 0.98m : 0.50m, agreement.ReviewStatus);
        AddSource(result, "FlexibleBenefitAgreement", agreement.SourceKey, "AdministratorSignaturePresent", officerInk?.Shape, BoolText(agreement.AdministratorSignaturePresent), officerSignatureLine, officerInk, officerInk?.Confidence ?? 0.45m, agreement.ReviewStatus);
    }

    private static FlexibleBenefitOcrLine? FindExplicitSignerLine(FlexibleBenefitOcrPage? page, string label)
    {
        // Accept one explicitly labelled name. Ambiguous or unlabelled names need human review.
        var candidates = page?.Lines.Where(line =>
                line.Text.TrimStart().StartsWith(label + ":", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(ValueAfterLabel(line.Text, label)))
            .Take(2).ToArray();
        return candidates is { Length: 1 } ? candidates[0] : null;
    }

    private void ExtractContacts(FlexibleBenefitExtractionResult result)
    {
        var document = result.OcrDocument!;
        var adminLabel = _layout.FindLine(document, "Company Plan Administrator");
        var adminName = adminLabel is null ? null : FindValueRightOfLabel(adminLabel, text => !Contains(text, "Company Plan Administrator"));
        var adminEmailLine = _layout.FindLine(document, "Email");
        var adminEmailRaw = adminEmailLine is null ? "" : ValueAfterLabel(adminEmailLine.Line.Text, "Email");
        var adminEmail = _normalizer.IsValidEmail(adminEmailRaw) ? adminEmailRaw : "";
        var planAdmin = new FlexibleBenefitContact
        {
            ContactRole = "PlanAdministrator",
            ContactName = adminName?.Line.Text ?? "",
            Email = adminEmail,
            Phone = result.Client.Phone,
            IsPrimary = true,
            SourcePageNumber = adminLabel?.Page.PageNumber,
            ExtractionConfidence = string.IsNullOrWhiteSpace(adminEmail) ? 0.62m : 0.91m,
            ReviewStatus = string.IsNullOrWhiteSpace(adminEmail)
                ? FlexibleBenefitReviewStatuses.NeedsReview
                : FlexibleBenefitReviewStatuses.Pending
        };
        result.Contacts.Add(planAdmin);
        AddSource(result, "FlexibleBenefitContact", planAdmin.SourceKey, "ContactName", adminName?.Line.Text, planAdmin.ContactName, adminName, null, 0.92m);
        AddSource(result, "FlexibleBenefitContact", planAdmin.SourceKey, "Email", adminEmailRaw, planAdmin.Email, adminEmailLine, null, planAdmin.ExtractionConfidence ?? 0.45m, planAdmin.ReviewStatus);
        if (string.IsNullOrWhiteSpace(adminEmail) && !string.IsNullOrWhiteSpace(adminEmailRaw))
        {
            result.ValidationIssues.Add(new FlexibleBenefitValidationIssue
            {
                Code = "invalid_plan_administrator_email",
                Severity = "Warning",
                FieldName = "PlanAdministrator.Email",
                Message = "The plan administrator email field is populated but is not a valid email address.",
                SourcePageNumber = adminEmailLine?.Page.PageNumber
            });
        }

        var brokerNameLabel = _layout.FindLine(document, "Broker Name");
        var brokerName = brokerNameLabel is null ? null : FindValueRightOfLabel(brokerNameLabel, text => !Contains(text, "Broker Name"));
        var brokerEmailLabel = _layout.FindLine(document, "Broker Email");
        var brokerEmail = brokerEmailLabel is null ? null : FindValueRightOfLabel(brokerEmailLabel, _normalizer.IsValidEmail);
        var broker = new FlexibleBenefitContact
        {
            ContactRole = "Broker",
            ContactName = brokerName?.Line.Text ?? "",
            Email = brokerEmail?.Line.Text ?? "",
            IsPrimary = false,
            SourcePageNumber = brokerNameLabel?.Page.PageNumber,
            ExtractionConfidence = 0.94m,
            ReviewStatus = FlexibleBenefitReviewStatuses.Pending
        };
        result.Contacts.Add(broker);
        AddSource(result, "FlexibleBenefitContact", broker.SourceKey, "ContactName", brokerName?.Line.Text, broker.ContactName, brokerName, null, 0.94m);
        AddSource(result, "FlexibleBenefitContact", broker.SourceKey, "Email", brokerEmail?.Line.Text, broker.Email, brokerEmail, null, 0.94m);

        var authorizationPage = document.Pages.FirstOrDefault(page => Contains(page.Text, "Signature of Authorized") && Contains(page.Text, "Company Officer"));
        var authorizationLine = authorizationPage is null
            ? null
            : _layout.FindLines(authorizationPage, "Print Name").OrderByDescending(match => match.MatchScore).FirstOrDefault();
        var officerName = authorizationLine is null ? null : FindValueRightOfLabel(authorizationLine, text => text.Length >= 4);
        if (officerName is not null)
        {
            var officer = new FlexibleBenefitContact
            {
                ContactRole = "AuthorizedOfficer",
                ContactName = officerName.Line.Text,
                IsPrimary = false,
                SourcePageNumber = officerName.Page.PageNumber,
                ExtractionConfidence = 0.94m,
                ReviewStatus = FlexibleBenefitReviewStatuses.Pending
            };
            result.Contacts.Add(officer);
            AddSource(result, "FlexibleBenefitContact", officer.SourceKey, "ContactName", officerName.Line.Text, officer.ContactName, officerName, null, 0.94m);
        }
    }

    private void ApplyPlanRollups(FlexibleBenefitExtractionResult result)
    {
        var plan = result.Plan;
        var firstClass = result.Classes.FirstOrDefault();
        var funding = result.FundingConfigurations.FirstOrDefault();
        var setupFee = result.Fees.FirstOrDefault(fee => fee.FeeType == "SetupFee");
        var adminFee = result.Fees.FirstOrDefault(fee => fee.FeeType == "AdministrationFee");
        plan.DefaultWaitingPeriodDays = firstClass?.WaitingPeriodDays;
        plan.NewHireProrationDefault = firstClass?.NewHireProrationEnabled;
        plan.FundingMethod = funding?.FundingMethod ?? FlexibleBenefitFundingMethod.Unknown;
        plan.PreFunded = funding?.PreFundingRequired;
        plan.PayAsYouGo = result.OcrDocument is not null && Contains(result.OcrDocument.CombinedText, "Pay-as-you-go funding is also available") ? true : null;
        plan.PadRequested = funding?.PadSelected;
        plan.SetupFee = setupFee?.FlatAmount;
        plan.AdministrationFeePercentage = adminFee?.Percentage;
        plan.AdministrationFeeBasis = adminFee?.Description ?? "";
        plan.ExtractionStatus = FlexibleBenefitExtractionStatuses.Normalized;
    }

    private async Task<(int? Value, string RawValue, FlexibleBenefitLineMatch? Match, decimal Confidence)> ExtractIntegerRightOfLabelAsync(
        FlexibleBenefitOcrDocument document,
        FlexibleBenefitLineMatch? label,
        CancellationToken cancellationToken)
    {
        if (label is null)
        {
            return (null, "", null, 0.35m);
        }

        var page = label.Page;
        var existing = page.Lines
            .Where(line => line.Bounds.X > label.Line.Bounds.Right)
            .Where(line => Math.Abs(line.Bounds.Y - label.Line.Bounds.Y) < page.PixelHeight * 0.025)
            .Select(line => new { Line = line, Value = _normalizer.ParseDependentAge(line.Text) })
            .FirstOrDefault(item => item.Value.HasValue);
        if (existing is not null)
        {
            return (existing.Value, existing.Line.Text, Match(page, existing.Line, label.Alias), 0.92m);
        }

        var x = label.Line.Bounds.Right + page.PixelWidth * 0.015;
        var region = new FlexibleBenefitBounds(
            x,
            Math.Max(0d, label.Line.Bounds.Y - page.PixelHeight * 0.015),
            Math.Min(page.PixelWidth * 0.28, page.PixelWidth - x),
            Math.Max(label.Line.Bounds.Height * 2.2, page.PixelHeight * 0.045));
        var regionLines = await _textExtractor.ExtractRegionAsync(
            document.FilePath,
            page.PageNumber,
            region,
            page.PixelWidth,
            page.PixelHeight,
            cancellationToken);
        var candidate = regionLines
            .Select(line => new { Line = line, Value = _normalizer.ParseDependentAge(line.Text) })
            .FirstOrDefault(item => item.Value.HasValue);
        return candidate is null
            ? (null, "", label, 0.45m)
            : (candidate.Value, candidate.Line.Text, Match(page, candidate.Line, label.Alias), 0.90m);
    }

    private FlexibleBenefitLineMatch? FindValueRightOfLabel(
        FlexibleBenefitLineMatch label,
        Func<string, bool> predicate)
    {
        var page = label.Page;
        var labelCenter = label.Line.Bounds.Y + label.Line.Bounds.Height / 2d;
        return page.Lines
            .Where(line => !ReferenceEquals(line, label.Line))
            .Where(line => line.Bounds.X >= label.Line.Bounds.Right - page.PixelWidth * 0.02)
            .Where(line => predicate(line.Text))
            .Where(line => Math.Abs(line.Bounds.Y + line.Bounds.Height / 2d - labelCenter) < page.PixelHeight * 0.04)
            .OrderBy(line => Math.Abs(line.Bounds.Y + line.Bounds.Height / 2d - labelCenter) + Math.Max(0d, line.Bounds.X - label.Line.Bounds.Right))
            .Select(line => new FlexibleBenefitLineMatch(page, line, label.Alias, 0.90m))
            .FirstOrDefault();
    }

    private FlexibleBenefitLineMatch? BestLine(FlexibleBenefitOcrPage page, params string[] aliases)
    {
        return _layout.FindLines(page, aliases)
            .OrderByDescending(match => match.MatchScore)
            .ThenBy(match => match.Line.Bounds.Y)
            .FirstOrDefault();
    }

    private static FlexibleBenefitLineMatch? FindHeader(FlexibleBenefitOcrPage page, string text)
    {
        var line = page.Lines
            .Where(line => Contains(line.Text, text))
            .Where(line => line.Bounds.Y < page.PixelHeight * 0.30)
            .OrderBy(line => line.Bounds.X)
            .FirstOrDefault();
        return line is null ? null : Match(page, line, text);
    }

    private FlexibleBenefitOcrLine? ClosestColumnValue(IReadOnlyList<FlexibleBenefitOcrLine> row, double columnX)
    {
        return row
            .Where(line => _normalizer.ParseMoney(line.Text).HasValue)
            .OrderBy(line => Math.Abs(line.Bounds.X - columnX))
            .FirstOrDefault();
    }

    private static FlexibleBenefitDetectedMark? FindMarkForLine(FlexibleBenefitOcrPage page, FlexibleBenefitOcrLine line)
    {
        var centerY = line.Bounds.Y + line.Bounds.Height / 2d;
        return page.Marks
            .Where(mark => mark.Shape.Equals("Selection mark", StringComparison.OrdinalIgnoreCase))
            .Where(mark => mark.Bounds.X <= line.Bounds.X + line.Bounds.Height)
            .Where(mark => line.Bounds.X - mark.Bounds.Right < page.PixelWidth * 0.10)
            .Where(mark => Math.Abs(mark.Bounds.Y + mark.Bounds.Height / 2d - centerY) < page.PixelHeight * 0.035)
            .OrderBy(mark => Math.Abs(line.Bounds.X - mark.Bounds.Right) + Math.Abs(mark.Bounds.Y + mark.Bounds.Height / 2d - centerY))
            .FirstOrDefault();
    }

    private static FlexibleBenefitDetectedMark? FindSignatureMark(
        FlexibleBenefitOcrPage? page,
        FlexibleBenefitLineMatch? signatureLine)
    {
        if (page is null || signatureLine is null)
        {
            return null;
        }

        return page.Marks
            .Where(mark => mark.Shape.Equals("Signature ink", StringComparison.OrdinalIgnoreCase))
            .OrderBy(mark => Math.Abs(mark.Bounds.Bottom - signatureLine.Line.Bounds.Y))
            .FirstOrDefault();
    }

    private void AddSource(
        FlexibleBenefitExtractionResult result,
        string entityType,
        string entityKey,
        string fieldName,
        string? rawValue,
        string? normalizedValue,
        FlexibleBenefitLineMatch? line,
        FlexibleBenefitDetectedMark? mark,
        decimal confidence,
        string reviewStatus = FlexibleBenefitReviewStatuses.Pending)
    {
        if (string.IsNullOrWhiteSpace(rawValue) && string.IsNullOrWhiteSpace(normalizedValue))
        {
            return;
        }

        result.SourceFields.Add(_layout.BuildSourceField(
            entityType,
            entityKey,
            fieldName,
            rawValue,
            normalizedValue,
            line,
            mark,
            result.OcrDocument?.ExtractionMethod ?? "Local OCR",
            confidence,
            reviewStatus));
    }

    private DateOnly? ParseAgreementDate(string value)
    {
        var normalized = Regex.Replace(value, @"20\s+(?<suffix>\d{2})", "20${suffix}");
        var month = Regex.Match(normalized, @"\b(January|February|March|April|May|June|July|August|September|October|November|December)\b", RegexOptions.IgnoreCase).Value;
        var day = Regex.Match(normalized, @"\b(?:the\s+)?(?<day>\d{1,2})\b", RegexOptions.IgnoreCase).Groups["day"].Value;
        var year = Regex.Match(normalized, @"\b20\d{2}\b").Value;
        return _normalizer.ParseDate($"{month} {day} {year}");
    }

    private static FlexibleBenefitLineMatch? Match(FlexibleBenefitOcrPage page, FlexibleBenefitOcrLine? line, string alias)
    {
        return line is null ? null : new FlexibleBenefitLineMatch(page, line, alias, 0.90m);
    }

    private static bool LooksLikeAddress(string text)
    {
        return Regex.IsMatch(text, @"\d+") && text.Any(char.IsLetter);
    }

    private static bool LooksLikeProvince(string text)
    {
        var normalized = text.Trim().Trim('.', ',');
        return normalized.Length == 2 && normalized.All(char.IsLetter) ||
               new[] { "Alberta", "British Columbia", "Manitoba", "Ontario", "Quebec", "Saskatchewan" }
                   .Contains(normalized, StringComparer.OrdinalIgnoreCase);
    }

    private static string ValueAfterLabel(string text, string label)
    {
        var index = text.IndexOf(label, StringComparison.OrdinalIgnoreCase);
        if (index < 0)
        {
            return "";
        }

        return text[(index + label.Length)..].Trim().TrimStart(':').Trim();
    }

    private static bool Contains(string? text, string value)
    {
        return (text ?? "").Contains(value, StringComparison.OrdinalIgnoreCase);
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }

    private static string BoolText(bool? value)
    {
        return value.HasValue ? (value.Value ? "true" : "false") : "";
    }

    private static string DecimalText(decimal? value)
    {
        return value?.ToString(CultureInfo.InvariantCulture) ?? "";
    }

    private static string IntegerText(int? value)
    {
        return value?.ToString(CultureInfo.InvariantCulture) ?? "";
    }

    private static string DateText(DateOnly? value)
    {
        return value?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "";
    }
}
