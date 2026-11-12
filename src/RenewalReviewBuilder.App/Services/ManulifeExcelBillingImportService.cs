using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ManulifeExcelBillingImportService
{
    private static readonly HashSet<string> KnownBenefitCodes = new(StringComparer.OrdinalIgnoreCase)
    {
        "ELI", "DLI", "OEL", "ADD", "EHC", "DEN", "LTD", "EAP", "EAT", "CPD", "CPH"
    };

    private readonly DatabaseService _databaseService;
    private readonly ManulifeExcelBillingImportOptions _options;

    public ManulifeExcelBillingImportService(
        DatabaseService? databaseService = null,
        ManulifeExcelBillingImportOptions? options = null)
    {
        _databaseService = databaseService ?? new DatabaseService();
        _options = options ?? new ManulifeExcelBillingImportOptions();
    }

    public static bool IsSupportedFile(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".csv", StringComparison.OrdinalIgnoreCase);
    }

    public static bool LooksLikeManulifeBillingWorkbook(string path)
    {
        if (!IsSupportedFile(path) || !File.Exists(path))
        {
            return false;
        }

        try
        {
            var table = BillingTableReader.Read(path, 2 * 1024 * 1024);
            var headers = table.Headers.Select(value => NormalizeHeader(value.Value)).ToHashSet(StringComparer.OrdinalIgnoreCase);
            return headers.Contains(NormalizeHeader("Certificate Num")) &&
                   headers.Contains(NormalizeHeader("All Benefits - Total Premium & Tax (4)")) &&
                   table.Headers.Any(header => BenefitHeaderRegex.IsMatch(header.Value ?? ""));
        }
        catch
        {
            return false;
        }
    }

    public async Task<ManulifeExcelBillingImportPreview> CreatePreviewAsync(ManulifeExcelBillingImportRequest request)
    {
        var preview = new ManulifeExcelBillingImportPreview
        {
            SelectedClientId = request.SelectedClientId,
            ClientName = Clean(request.SelectedClientName),
            SourceFilePath = request.InputPath,
            SourceFileName = Path.GetFileName(request.InputPath),
            StatementDate = request.StatementDate
        };

        if (string.IsNullOrWhiteSpace(request.SelectedClientName) && !request.SelectedClientId.HasValue)
        {
            preview.RequiresClientConfirmation = true;
            AddException(preview, null, "", "", "client", "client_confirmation_required", CarrierBillingExceptionSeverities.Warning,
                "", "", "Select the client before confirming this billing import.");
        }

        if (!IsSupportedFile(request.InputPath))
        {
            AddException(preview, null, "", "", "file", "unsupported_file_type", CarrierBillingExceptionSeverities.Error,
                Path.GetExtension(request.InputPath), "xlsx, xlsm, or csv", "Only .xlsx, .xlsm, and .csv Manulife billing files are supported.");
            return preview;
        }

        if (!File.Exists(request.InputPath))
        {
            AddException(preview, null, "", "", "file", "file_not_found", CarrierBillingExceptionSeverities.Error,
                "", "", "The billing source file was not found.");
            return preview;
        }

        var fileInfo = new FileInfo(request.InputPath);
        if (fileInfo.Length > _options.MaximumFileSizeBytes)
        {
            AddException(preview, null, "", "", "file", "file_too_large", CarrierBillingExceptionSeverities.Error,
                fileInfo.Length.ToString(CultureInfo.InvariantCulture), _options.MaximumFileSizeBytes.ToString(CultureInfo.InvariantCulture), "The billing file exceeds the configured size limit.");
            return preview;
        }

        try
        {
            preview.SourceFileHash = await ComputeFileHashAsync(request.InputPath);
            preview.BillingPeriod = FirstNonBlank(NormalizeBillingPeriod(request.BillingPeriod), InferBillingPeriodFromFileName(request.InputPath));
            if (string.IsNullOrWhiteSpace(preview.BillingPeriod))
            {
                AddException(preview, null, "", "", "billing_period", "billing_period_required", CarrierBillingExceptionSeverities.Error,
                    "", "", "Confirm a billing period before importing this file.");
            }

            if (request.SelectedClientId.HasValue)
            {
                preview.DuplicateFileDetected = await _databaseService.IsCarrierBillingFileAlreadyImportedAsync(
                    request.SelectedClientId.Value,
                    preview.CarrierName,
                    preview.SourceFileHash);
                if (preview.DuplicateFileDetected)
                {
                    AddException(preview, null, "", "", "source_file_hash", "duplicate_source_file", CarrierBillingExceptionSeverities.Error,
                        "", "", "This source file has already been imported for the selected client.");
                }
            }

            var table = BillingTableReader.Read(request.InputPath, _options.MaximumFileSizeBytes);
            preview.SourceRowCount = table.Rows.Count;
            var headerMap = ManulifeExcelBillingHeaderMapper.Map(table.Headers);
            preview.HeaderMappings.AddRange(headerMap.Mappings);
            preview.NewBenefitCodes.AddRange(headerMap.BenefitColumns.Keys
                .Where(code => !KnownBenefitCodes.Contains(code))
                .OrderBy(code => code, StringComparer.OrdinalIgnoreCase));
            preview.UnmappedHeaders.AddRange(headerMap.Mappings
                .Where(mapping => mapping.Status == CarrierBillingHeaderStatuses.Unmapped)
                .Select(mapping => mapping.SourceHeader));

            foreach (var mapping in headerMap.Mappings.Where(mapping => mapping.Status == CarrierBillingHeaderStatuses.Ambiguous))
            {
                AddException(preview, null, "", mapping.BenefitCode, mapping.TargetField, "ambiguous_header", CarrierBillingExceptionSeverities.Error,
                    mapping.SourceHeader, "", "A required header was mapped more than once. Resolve the header mapping before import.");
            }

            foreach (var required in ManulifeExcelBillingHeaderMapper.RequiredCoreFields.Where(required => !headerMap.CoreColumns.ContainsKey(required)))
            {
                AddException(preview, null, "", "", required, "required_header_missing", CarrierBillingExceptionSeverities.Error,
                    "", required, "A required Manulife billing header was not found.");
            }

            foreach (var newBenefit in preview.NewBenefitCodes)
            {
                AddException(preview, null, "", newBenefit, "benefit_code", "new_benefit_code", CarrierBillingExceptionSeverities.Warning,
                    newBenefit, "", "A new benefit code was detected and will be retained in the import.");
            }

            if (preview.ErrorCount > 0)
            {
                return preview;
            }

            ParseRows(preview, table, headerMap);
            preview.DetectedClientName = InferClientNameFromClassNames(preview.EmployeeRecords);
            Reconcile(preview);
            BuildAggregates(preview);
            BuildPreviewSample(preview);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            AddException(preview, null, "", "", "file", "billing_file_read_error", CarrierBillingExceptionSeverities.Error,
                "", "", "The Manulife billing file could not be read. No employee data was saved.");
        }

        return preview;
    }

    public Task<CarrierBillingConfirmationResult> ConfirmAsync(
        ManulifeExcelBillingImportPreview preview,
        ManulifeExcelBillingImportRequest request)
    {
        return _databaseService.SaveManulifeExcelBillingImportAsync(preview, request);
    }

    private void ParseRows(ManulifeExcelBillingImportPreview preview, BillingTable table, ManulifeExcelBillingHeaderMap headerMap)
    {
        foreach (var row in table.Rows)
        {
            if (!row.Cells.Any(cell => !string.IsNullOrWhiteSpace(cell.Value)))
            {
                continue;
            }

            var certificate = ReadText(row, headerMap, "certificate_number");
            if (string.IsNullOrWhiteSpace(certificate))
            {
                AddException(preview, row.RowNumber, "", "", "certificate_number", "certificate_number_missing", CarrierBillingExceptionSeverities.Error,
                    "", "", "A populated billing row did not include a certificate number.");
                continue;
            }

            ValidateFormulaCells(preview, row);
            var employee = new CarrierBillingEmployeeRecord
            {
                CertificateNumber = certificate,
                LastName = ReadText(row, headerMap, "last_name"),
                FirstName = ReadText(row, headerMap, "first_name"),
                ClassName = ReadText(row, headerMap, "class_name"),
                ClassCode = ReadText(row, headerMap, "class_code"),
                PlanName = ReadText(row, headerMap, "plan_name"),
                PlanCode = ReadText(row, headerMap, "plan_code"),
                ProvinceOfWorkRaw = ReadText(row, headerMap, "province_of_work"),
                ProvinceOfResidenceRaw = ReadText(row, headerMap, "province_of_residence"),
                CurrentMonthPremium = ReadMoney(row, headerMap, "current_month_premium", preview),
                PriorPeriodAdjustment = ReadMoney(row, headerMap, "prior_period_adjustment", preview),
                TotalSalesTax = ReadMoney(row, headerMap, "total_sales_tax", preview),
                TotalPremiumAndTax = ReadMoney(row, headerMap, "total_premium_and_tax", preview),
                DivisionalBillingCode = ReadText(row, headerMap, "divisional_billing_code"),
                BillTotal6 = ReadText(row, headerMap, "bill_total_6"),
                BillTotal7 = ReadText(row, headerMap, "bill_total_7"),
                SourceRowNumber = row.RowNumber,
                SourceRowHash = ComputeTextHash(string.Join("|", row.Cells.Select(cell => cell.Value))),
                ValidationStatus = "Passed"
            };
            employee.ProvinceOfWork = NormalizeProvince(employee.ProvinceOfWorkRaw);
            employee.ProvinceOfResidence = NormalizeProvince(employee.ProvinceOfResidenceRaw);
            preview.EmployeeRecords.Add(employee);

            foreach (var benefitGroup in headerMap.BenefitColumns.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase))
            {
                var benefit = ReadBenefit(row, employee, benefitGroup.Key, benefitGroup.Value, preview);
                if (benefit is not null)
                {
                    preview.BenefitRecords.Add(benefit);
                }
            }
        }

        preview.EmployeeRecordCount = preview.EmployeeRecords.Count;
        preview.BenefitRecordCount = preview.BenefitRecords.Count;
    }

    private CarrierBillingBenefitRecord? ReadBenefit(
        BillingTableRow row,
        CarrierBillingEmployeeRecord employee,
        string benefitCode,
        IReadOnlyDictionary<string, int> columns,
        ManulifeExcelBillingImportPreview preview)
    {
        var planCode = ReadText(row, columns, "plan_code");
        var rawDependentOrVolume = ReadText(row, columns, "dependent_code_or_volume");
        var option = ReadText(row, columns, "option");
        var current = ReadMoney(row, columns, "current_month_premium", preview, benefitCode);
        var adjustment = ReadMoney(row, columns, "prior_period_adjustment", preview, benefitCode);
        var tax = ReadMoney(row, columns, "total_sales_tax", preview, benefitCode);
        var total = ReadMoney(row, columns, "benefit_total", preview, benefitCode);
        var hasCoverage = !string.IsNullOrWhiteSpace(planCode) || !string.IsNullOrWhiteSpace(rawDependentOrVolume) || !string.IsNullOrWhiteSpace(option);
        var hasFinancialValue = current.HasValue || adjustment.HasValue || tax.HasValue || total.HasValue;
        if (!hasCoverage && !hasFinancialValue)
        {
            return null;
        }

        var benefit = new CarrierBillingBenefitRecord
        {
            CertificateNumber = employee.CertificateNumber,
            BenefitCode = benefitCode,
            BenefitName = BenefitName(benefitCode),
            BenefitPlanCode = planCode,
            DependentCodeOrVolumeRaw = rawDependentOrVolume,
            DependentCode = IsDependentCode(rawDependentOrVolume) ? rawDependentOrVolume : "",
            InsuredVolume = IsDependentCode(rawDependentOrVolume) ? null : ParseMoney(rawDependentOrVolume),
            OptionCode = option,
            CurrentMonthPremium = current,
            PriorPeriodAdjustment = adjustment,
            TotalSalesTax = tax,
            BenefitTotal = total,
            CalculatedBenefitTotal = RoundMoney((current ?? 0m) + (adjustment ?? 0m) + (tax ?? 0m)),
            SourceRowNumber = row.RowNumber,
            ValidationStatus = "Passed"
        };
        if (total.HasValue)
        {
            benefit.ReconciliationDifference = RoundMoney(total.Value - benefit.CalculatedBenefitTotal);
            ApplyVariance(preview, employee, benefit, "benefit_total", benefit.ReconciliationDifference.Value, total.Value, benefit.CalculatedBenefitTotal);
        }

        return benefit;
    }

    private void Reconcile(ManulifeExcelBillingImportPreview preview)
    {
        foreach (var employee in preview.EmployeeRecords)
        {
            var benefits = preview.BenefitRecords.Where(record => record.SourceRowNumber == employee.SourceRowNumber).ToList();
            employee.CalculatedCurrentMonthPremium = RoundMoney(benefits.Sum(record => record.CurrentMonthPremium ?? 0m));
            employee.CalculatedPriorPeriodAdjustment = RoundMoney(benefits.Sum(record => record.PriorPeriodAdjustment ?? 0m));
            employee.CalculatedSalesTax = RoundMoney(benefits.Sum(record => record.TotalSalesTax ?? 0m));
            employee.CalculatedPremiumAndTax = RoundMoney(benefits.Sum(record => record.BenefitTotal ?? record.CalculatedBenefitTotal));
            ReconcileEmployeeField(preview, employee, "current_month_premium", employee.CurrentMonthPremium, employee.CalculatedCurrentMonthPremium);
            ReconcileEmployeeField(preview, employee, "prior_period_adjustment", employee.PriorPeriodAdjustment, employee.CalculatedPriorPeriodAdjustment);
            ReconcileEmployeeField(preview, employee, "total_sales_tax", employee.TotalSalesTax, employee.CalculatedSalesTax);
            ReconcileEmployeeField(preview, employee, "total_premium_and_tax", employee.TotalPremiumAndTax, employee.CalculatedPremiumAndTax);
        }

        foreach (var group in preview.EmployeeRecords.GroupBy(record => record.CertificateNumber, StringComparer.OrdinalIgnoreCase).Where(group => group.Count() > 1))
        {
            foreach (var employee in group)
            {
                employee.ValidationStatus = "Warning";
                AddException(preview, employee.SourceRowNumber, employee.CertificateNumber, "", "certificate_number", "duplicate_certificate_in_file", CarrierBillingExceptionSeverities.Warning,
                    "", "", "The certificate number occurs more than once in this billing file.");
            }
        }

        foreach (var group in preview.EmployeeRecords.GroupBy(record => record.SourceRowHash, StringComparer.OrdinalIgnoreCase).Where(group => group.Count() > 1))
        {
            foreach (var employee in group)
            {
                employee.ValidationStatus = "Warning";
                AddException(preview, employee.SourceRowNumber, employee.CertificateNumber, "", "source_row", "duplicate_source_row", CarrierBillingExceptionSeverities.Warning,
                    "", "", "An exact duplicate employee row was detected in this billing file.");
            }
        }

        preview.TotalCurrentMonthPremium = RoundMoney(preview.EmployeeRecords.Sum(record => record.CurrentMonthPremium ?? 0m));
        preview.TotalPriorPeriodAdjustment = RoundMoney(preview.EmployeeRecords.Sum(record => record.PriorPeriodAdjustment ?? 0m));
        preview.TotalSalesTax = RoundMoney(preview.EmployeeRecords.Sum(record => record.TotalSalesTax ?? 0m));
        preview.TotalPremiumAndTax = RoundMoney(preview.EmployeeRecords.Sum(record => record.TotalPremiumAndTax ?? 0m));
        preview.CalculatedPremiumAndTax = RoundMoney(preview.EmployeeRecords.Sum(record => record.CalculatedPremiumAndTax));
        preview.ReconciliationDifference = RoundMoney(preview.TotalPremiumAndTax - preview.CalculatedPremiumAndTax);
        if (preview.ReconciliationDifference != 0m)
        {
            AddException(preview, null, "", "", "file_total", "file_reconciliation_difference",
                Math.Abs(preview.ReconciliationDifference) <= _options.MinorVarianceThreshold ? CarrierBillingExceptionSeverities.Warning : CarrierBillingExceptionSeverities.Error,
                preview.TotalPremiumAndTax.ToString("0.00", CultureInfo.InvariantCulture), preview.CalculatedPremiumAndTax.ToString("0.00", CultureInfo.InvariantCulture), "The file-level employee total does not match the calculated benefit total.");
        }
    }

    private void ReconcileEmployeeField(ManulifeExcelBillingImportPreview preview, CarrierBillingEmployeeRecord employee, string field, decimal? reported, decimal calculated)
    {
        if (!reported.HasValue)
        {
            return;
        }

        var difference = RoundMoney(reported.Value - calculated);
        if (field == "total_premium_and_tax")
        {
            employee.ReconciliationDifference = difference;
        }

        ApplyVariance(preview, employee, null, field, difference, reported.Value, calculated);
    }

    private void ApplyVariance(ManulifeExcelBillingImportPreview preview, CarrierBillingEmployeeRecord employee, CarrierBillingBenefitRecord? benefit, string field, decimal difference, decimal reported, decimal calculated)
    {
        if (difference == 0m)
        {
            return;
        }

        var severity = Math.Abs(difference) <= _options.MinorVarianceThreshold
            ? CarrierBillingExceptionSeverities.Warning
            : CarrierBillingExceptionSeverities.Error;
        employee.ValidationStatus = severity == CarrierBillingExceptionSeverities.Error ? "Failed" : "Warning";
        if (benefit is not null)
        {
            benefit.ValidationStatus = employee.ValidationStatus;
        }

        AddException(preview, employee.SourceRowNumber, employee.CertificateNumber, benefit?.BenefitCode ?? "", field,
            benefit is null ? "employee_benefit_reconciliation_difference" : "benefit_reconciliation_difference",
            severity,
            reported.ToString("0.00", CultureInfo.InvariantCulture), calculated.ToString("0.00", CultureInfo.InvariantCulture),
            "The carrier-reported amount does not match the calculated benefit amount.");
    }

    private static void BuildAggregates(ManulifeExcelBillingImportPreview preview)
    {
        AddAggregates(preview.TotalsByBenefit, preview.BenefitRecords, record => record.BenefitCode);
        AddAggregates(preview.TotalsByClass, preview.EmployeeRecords, record => record.ClassCode);
        AddAggregates(preview.TotalsByPlan, preview.EmployeeRecords, record => record.PlanCode);
        AddAggregates(preview.TotalsByDivision, preview.EmployeeRecords, record => record.DivisionalBillingCode);
    }

    private static void AddAggregates<T>(ICollection<CarrierBillingAggregate> target, IEnumerable<T> records, Func<T, string> groupKey) where T : class
    {
        foreach (var group in records.GroupBy(groupKey).OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            var values = group.Cast<object>().ToList();
            var aggregate = new CarrierBillingAggregate
            {
                Group = FirstNonBlank(group.Key, "Not specified"),
                EmployeeCount = values.OfType<CarrierBillingEmployeeRecord>().Select(record => record.CertificateNumber).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                BenefitRecordCount = values.OfType<CarrierBillingBenefitRecord>().Count(),
                CurrentMonthPremium = RoundMoney(values.Sum(GetCurrentPremium)),
                PriorPeriodAdjustment = RoundMoney(values.Sum(GetAdjustment)),
                SalesTax = RoundMoney(values.Sum(GetTax)),
                TotalPremiumAndTax = RoundMoney(values.Sum(GetTotal))
            };
            target.Add(aggregate);
        }
    }

    private static decimal GetCurrentPremium(object value) => value switch
    {
        CarrierBillingEmployeeRecord employee => employee.CurrentMonthPremium ?? 0m,
        CarrierBillingBenefitRecord benefit => benefit.CurrentMonthPremium ?? 0m,
        _ => 0m
    };

    private static decimal GetAdjustment(object value) => value switch
    {
        CarrierBillingEmployeeRecord employee => employee.PriorPeriodAdjustment ?? 0m,
        CarrierBillingBenefitRecord benefit => benefit.PriorPeriodAdjustment ?? 0m,
        _ => 0m
    };

    private static decimal GetTax(object value) => value switch
    {
        CarrierBillingEmployeeRecord employee => employee.TotalSalesTax ?? 0m,
        CarrierBillingBenefitRecord benefit => benefit.TotalSalesTax ?? 0m,
        _ => 0m
    };

    private static decimal GetTotal(object value) => value switch
    {
        CarrierBillingEmployeeRecord employee => employee.TotalPremiumAndTax ?? 0m,
        CarrierBillingBenefitRecord benefit => benefit.BenefitTotal ?? benefit.CalculatedBenefitTotal,
        _ => 0m
    };

    private static void BuildPreviewSample(ManulifeExcelBillingImportPreview preview)
    {
        preview.SampleEmployeeRecords.AddRange(preview.EmployeeRecords.Take(10).Select(employee => new CarrierBillingPreviewEmployee
        {
            SourceRowNumber = employee.SourceRowNumber,
            MaskedCertificateNumber = Mask(employee.CertificateNumber),
            MaskedEmployeeName = MaskName(employee.FirstName, employee.LastName),
            ClassCode = employee.ClassCode,
            PlanCode = employee.PlanCode,
            TotalPremiumAndTax = employee.TotalPremiumAndTax,
            ValidationStatus = employee.ValidationStatus
        }));
    }

    private static void ValidateFormulaCells(ManulifeExcelBillingImportPreview preview, BillingTableRow row)
    {
        foreach (var cell in row.Cells.Where(cell => cell.IsFormula || LooksLikeFormulaInput(cell.Value)))
        {
            AddException(preview, row.RowNumber, "", "", "source_cell", "formula_or_csv_injection", CarrierBillingExceptionSeverities.Error,
                "", "", "Formula-like source content is not accepted in billing import data.");
        }
    }

    private static decimal? ReadMoney(BillingTableRow row, ManulifeExcelBillingHeaderMap map, string field, ManulifeExcelBillingImportPreview preview, string benefitCode = "")
    {
        return map.CoreColumns.TryGetValue(field, out var column)
            ? ReadMoney(row, column, preview, benefitCode, field)
            : null;
    }

    private static decimal? ReadMoney(BillingTableRow row, IReadOnlyDictionary<string, int> map, string field, ManulifeExcelBillingImportPreview preview, string benefitCode)
    {
        return map.TryGetValue(field, out var column)
            ? ReadMoney(row, column, preview, benefitCode, field)
            : null;
    }

    private static decimal? ReadMoney(BillingTableRow row, int column, ManulifeExcelBillingImportPreview preview, string benefitCode, string field)
    {
        var raw = ReadText(row, column);
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        var parsed = ParseMoney(raw);
        if (!parsed.HasValue)
        {
            AddException(preview, row.RowNumber, "", benefitCode, field, "invalid_money_value", CarrierBillingExceptionSeverities.Error,
                "", "decimal", "A monetary value could not be converted to a decimal amount.");
        }

        return parsed;
    }

    private static string ReadText(BillingTableRow row, ManulifeExcelBillingHeaderMap map, string field)
    {
        return map.CoreColumns.TryGetValue(field, out var column) ? ReadText(row, column) : "";
    }

    private static string ReadText(BillingTableRow row, IReadOnlyDictionary<string, int> map, string field)
    {
        return map.TryGetValue(field, out var column) ? ReadText(row, column) : "";
    }

    private static string ReadText(BillingTableRow row, int column)
    {
        return column >= 0 && column < row.Cells.Count ? Clean(row.Cells[column].Value) : "";
    }

    private static void AddException(ManulifeExcelBillingImportPreview preview, int? rowNumber, string certificate, string benefitCode, string fieldName, string type, string severity, string sourceValue, string expectedValue, string message)
    {
        preview.Exceptions.Add(new CarrierBillingException
        {
            SourceRowNumber = rowNumber,
            CertificateNumber = certificate,
            BenefitCode = benefitCode,
            FieldName = fieldName,
            ExceptionType = type,
            Severity = severity,
            SourceValue = sourceValue,
            ExpectedValue = expectedValue,
            Message = message
        });
    }

    private static string NormalizeProvince(string value)
    {
        var normalized = Clean(value).ToUpperInvariant();
        return normalized switch
        {
            "BRITISH COLUMBIA" or "BC" => "BC",
            "ALBERTA" or "AB" => "AB",
            "SASKATCHEWAN" or "SK" => "SK",
            "MANITOBA" or "MB" => "MB",
            "ONTARIO" or "ON" => "ON",
            "QUEBEC" or "QC" or "PQ" => "QC",
            "NEW BRUNSWICK" or "NB" => "NB",
            "NOVA SCOTIA" or "NS" => "NS",
            "PRINCE EDWARD ISLAND" or "PE" => "PE",
            "NEWFOUNDLAND" or "NEWFOUNDLAND AND LABRADOR" or "NL" => "NL",
            "NORTHWEST TERRITORIES" or "NT" => "NT",
            "NUNAVUT" or "NU" => "NU",
            "YUKON" or "YT" => "YT",
            _ => Clean(value)
        };
    }

    private static bool IsDependentCode(string value)
    {
        return value.Equals("Single", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Couple", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Family", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Employee", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Waived", StringComparison.OrdinalIgnoreCase);
    }

    private static decimal? ParseMoney(string? value)
    {
        var cleaned = Clean(value).Replace("$", "", StringComparison.Ordinal).Replace(",", "", StringComparison.Ordinal);
        if (cleaned.Length == 0)
        {
            return null;
        }

        var negative = cleaned.StartsWith("(", StringComparison.Ordinal) && cleaned.EndsWith(")", StringComparison.Ordinal);
        cleaned = cleaned.Trim('(', ')');
        return decimal.TryParse(cleaned, NumberStyles.Number | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var amount)
            ? negative ? -amount : amount
            : null;
    }

    private static bool LooksLikeFormulaInput(string value)
    {
        var trimmed = Clean(value);
        if (trimmed.Length == 0 || ParseMoney(trimmed).HasValue)
        {
            return false;
        }

        return trimmed.StartsWith("=", StringComparison.Ordinal) ||
               trimmed.StartsWith("+", StringComparison.Ordinal) ||
               trimmed.StartsWith("-", StringComparison.Ordinal) ||
               trimmed.StartsWith("@", StringComparison.Ordinal);
    }

    private static string BenefitName(string code) => code.ToUpperInvariant() switch
    {
        "ELI" => "Employee Life Insurance",
        "DLI" => "Dependent Life Insurance",
        "OEL" => "Optional Employee Life",
        "ADD" => "Accidental Death and Dismemberment",
        "EHC" => "Extended Health Care",
        "DEN" => "Dental",
        "LTD" => "Long-Term Disability",
        "EAP" => "Employee Assistance Program",
        "EAT" => "Employee Assistance Technology",
        _ => code
    };

    private static string InferBillingPeriodFromFileName(string path)
    {
        var match = Regex.Match(Path.GetFileNameWithoutExtension(path), @"(?<month>Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s*[-_ ]\s*(?<year>20\d{2})", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return "";
        }

        return DateTime.TryParseExact($"{match.Groups["month"].Value} {match.Groups["year"].Value}", new[] { "MMM yyyy", "MMMM yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
            ? date.ToString("yyyy-MM", CultureInfo.InvariantCulture)
            : "";
    }

    private static string InferClientNameFromClassNames(IEnumerable<CarrierBillingEmployeeRecord> employees)
    {
        var candidates = employees
            .Select(employee => RemoveClassCodePrefix(employee.ClassName))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .GroupBy(value => NormalizeClientCandidate(value), StringComparer.OrdinalIgnoreCase)
            .Select(group => new
            {
                DisplayName = group.First(),
                Count = group.Count()
            })
            .OrderByDescending(candidate => candidate.Count)
            .ThenBy(candidate => candidate.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return candidates.Count == 0 ? "" : candidates[0].DisplayName;
    }

    private static string RemoveClassCodePrefix(string value)
    {
        var cleaned = Clean(value);
        var match = Regex.Match(
            cleaned,
            @"^(?:class\s*)?[A-Za-z0-9]{1,12}\s*[-:]\s*(?<client>.+)$",
            RegexOptions.IgnoreCase);
        return match.Success ? Clean(match.Groups["client"].Value) : "";
    }

    private static string NormalizeClientCandidate(string value)
    {
        return Regex.Replace(Clean(value).ToUpperInvariant(), "[^A-Z0-9]+", " ").Trim();
    }

    private static string NormalizeBillingPeriod(string value)
    {
        var cleaned = Clean(value);
        if (DateTime.TryParse(cleaned, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var date))
        {
            return date.ToString("yyyy-MM", CultureInfo.InvariantCulture);
        }

        return Regex.IsMatch(cleaned, @"^20\d{2}-\d{2}$") ? cleaned : "";
    }

    private static string Mask(string value)
    {
        var cleaned = Clean(value);
        return cleaned.Length <= 2 ? new string('*', cleaned.Length) : new string('*', cleaned.Length - 2) + cleaned[^2..];
    }

    private static string MaskName(string firstName, string lastName)
    {
        return string.Join(" ", new[] { Clean(firstName), Clean(lastName) }
            .Where(value => value.Length > 0)
            .Select(value => value[0] + new string('*', Math.Max(0, value.Length - 1))));
    }

    private static decimal RoundMoney(decimal value) => Math.Round(value, 2, MidpointRounding.AwayFromZero);
    private static string Clean(string? value) => (value ?? "").Trim();
    private static string FirstNonBlank(params string[] values) => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";

    private static string ComputeTextHash(string value)
    {
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value ?? "")));
    }

    private static async Task<string> ComputeFileHashAsync(string path)
    {
        await using var stream = File.OpenRead(path);
        return Convert.ToHexString(await SHA256.HashDataAsync(stream));
    }

    private static readonly Regex BenefitHeaderRegex = new(@"^(?<code>[A-Za-z0-9]{2,12})\s*:\s*(?<field>.+)$", RegexOptions.Compiled);

    public static string NormalizeHeader(string? value)
    {
        var normalized = (value ?? "")
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Replace('’', '\'')
            .Replace('‘', '\'')
            .Replace("&", "and", StringComparison.OrdinalIgnoreCase);
        normalized = Regex.Replace(normalized, @"\s+", " ").Trim();
        normalized = normalized.Replace("'", "", StringComparison.Ordinal);
        normalized = Regex.Replace(normalized, @"[^A-Za-z0-9]+", "");
        return normalized.ToLowerInvariant();
    }

    private sealed class BillingTable
    {
        public List<BillingCell> Headers { get; } = [];
        public List<BillingTableRow> Rows { get; } = [];
    }

    private sealed class BillingTableRow
    {
        public int RowNumber { get; init; }
        public List<BillingCell> Cells { get; } = [];
    }

    private sealed class BillingCell
    {
        public string Value { get; init; } = "";
        public bool IsFormula { get; init; }
    }

    private sealed class ManulifeExcelBillingHeaderMap
    {
        public Dictionary<string, int> CoreColumns { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, Dictionary<string, int>> BenefitColumns { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<CarrierBillingHeaderMapping> Mappings { get; } = [];
    }

    private static class ManulifeExcelBillingHeaderMapper
    {
        public static readonly string[] RequiredCoreFields = ["certificate_number", "current_month_premium", "total_premium_and_tax"];

        private static readonly Dictionary<string, string> CanonicalHeaders = new(StringComparer.OrdinalIgnoreCase)
        {
            ["certificate_number"] = "Certificate Num",
            ["last_name"] = "Last Name",
            ["first_name"] = "First Name",
            ["class_name"] = "Class Name",
            ["class_code"] = "Class Code",
            ["plan_name"] = "Plan Name",
            ["plan_code"] = "Plan Code",
            ["province_of_work"] = "Province of Work",
            ["province_of_residence"] = "Province of Res",
            ["current_month_premium"] = "All Benefits - Current Month Premium (1)",
            ["prior_period_adjustment"] = "All Benefits - Prior Period Adjustments' Premium (2)",
            ["total_sales_tax"] = "All Benefits - Total Sales Tax (3)",
            ["total_premium_and_tax"] = "All Benefits - Total Premium & Tax (4)",
            ["divisional_billing_code"] = "Divisional Billing",
            ["bill_total_6"] = "Bill Total 6",
            ["bill_total_7"] = "Bill Total 7"
        };

        private static readonly Dictionary<string, string> BenefitFields = new(StringComparer.OrdinalIgnoreCase)
        {
            [NormalizeHeader("Plan Code")] = "plan_code",
            [NormalizeHeader("Dep Code or Volume")] = "dependent_code_or_volume",
            [NormalizeHeader("Option")] = "option",
            [NormalizeHeader("Current Month Premium (1)")] = "current_month_premium",
            [NormalizeHeader("Prior Period Adjustments' Premium (2)")] = "prior_period_adjustment",
            [NormalizeHeader("Total Sales Tax (3)")] = "total_sales_tax",
            [NormalizeHeader("Benefit Total (4)")] = "benefit_total"
        };

        private static readonly Dictionary<string, string[]> BuiltInAliases = new(StringComparer.OrdinalIgnoreCase)
        {
            ["certificate_number"] = ["Certificate Number", "Certificate #"],
            ["last_name"] = ["Surname"],
            ["first_name"] = ["Given Name"],
            ["class_code"] = ["Class"],
            ["plan_code"] = ["Plan"],
            ["province_of_work"] = ["Work Province"],
            ["province_of_residence"] = ["Province of Residence", "Residence Province"],
            ["current_month_premium"] = ["All Benefits Current Month Premium"],
            ["prior_period_adjustment"] = ["All Benefits - Prior Period Adjustments Premium (2)"],
            ["total_sales_tax"] = ["All Benefits Total Sales Tax"],
            ["total_premium_and_tax"] = ["All Benefits Total Premium and Tax"],
            ["divisional_billing_code"] = ["Division Billing"]
        };

        public static ManulifeExcelBillingHeaderMap Map(IReadOnlyList<BillingCell> headers)
        {
            var aliases = LoadAliases();
            var map = new ManulifeExcelBillingHeaderMap();
            var coreMatches = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

            for (var index = 0; index < headers.Count; index++)
            {
                var source = headers[index].Value;
                var normalized = NormalizeHeader(source);
                var mapping = new CarrierBillingHeaderMapping
                {
                    ColumnNumber = index + 1,
                    SourceHeader = source,
                    NormalizedHeader = normalized
                };
                var core = aliases.FirstOrDefault(pair => pair.Value.Contains(normalized, StringComparer.OrdinalIgnoreCase));
                if (!string.IsNullOrWhiteSpace(core.Key))
                {
                    mapping.TargetField = core.Key;
                    mapping.Status = normalized == NormalizeHeader(CanonicalHeaders[core.Key])
                        ? CarrierBillingHeaderStatuses.ExactMatch
                        : CarrierBillingHeaderStatuses.ApprovedAlias;
                    if (!coreMatches.TryGetValue(core.Key, out var matches))
                    {
                        matches = [];
                        coreMatches[core.Key] = matches;
                    }
                    matches.Add(index);
                    map.Mappings.Add(mapping);
                    continue;
                }

                var benefitMatch = BenefitHeaderRegex.Match(source ?? "");
                if (benefitMatch.Success)
                {
                    var code = benefitMatch.Groups["code"].Value.Trim().ToUpperInvariant();
                    var field = NormalizeHeader(benefitMatch.Groups["field"].Value);
                    mapping.BenefitCode = code;
                    if (BenefitFields.TryGetValue(field, out var target))
                    {
                        mapping.TargetField = target;
                        mapping.Status = CarrierBillingHeaderStatuses.DynamicallyDetectedBenefit;
                        if (!map.BenefitColumns.TryGetValue(code, out var group))
                        {
                            group = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                            map.BenefitColumns[code] = group;
                        }
                        if (group.ContainsKey(target))
                        {
                            mapping.Status = CarrierBillingHeaderStatuses.Ambiguous;
                        }
                        else
                        {
                            group[target] = index;
                        }
                    }
                    else
                    {
                        mapping.Status = CarrierBillingHeaderStatuses.Unmapped;
                    }
                }
                else
                {
                    mapping.Status = CarrierBillingHeaderStatuses.Unmapped;
                }

                map.Mappings.Add(mapping);
            }

            foreach (var match in coreMatches.Where(pair => pair.Value.Count > 1))
            {
                foreach (var index in match.Value)
                {
                    map.Mappings.First(mapping => mapping.ColumnNumber == index + 1).Status = CarrierBillingHeaderStatuses.Ambiguous;
                }
                map.CoreColumns.Remove(match.Key);
            }

            foreach (var match in coreMatches.Where(pair => pair.Value.Count == 1))
            {
                map.CoreColumns[match.Key] = match.Value[0];
            }

            foreach (var benefit in map.BenefitColumns.Values)
            {
                var duplicateFields = map.Mappings.Where(mapping => mapping.BenefitCode.Length > 0 && mapping.Status == CarrierBillingHeaderStatuses.Ambiguous)
                    .Select(mapping => mapping.TargetField)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                foreach (var field in duplicateFields)
                {
                    benefit.Remove(field);
                }
            }

            return map;
        }

        private static Dictionary<string, HashSet<string>> LoadAliases()
        {
            var aliases = CanonicalHeaders.ToDictionary(
                pair => pair.Key,
                pair => new HashSet<string>(new[] { NormalizeHeader(pair.Value) }, StringComparer.OrdinalIgnoreCase),
                StringComparer.OrdinalIgnoreCase);
            foreach (var entry in BuiltInAliases)
            {
                foreach (var alias in entry.Value)
                {
                    aliases[entry.Key].Add(NormalizeHeader(alias));
                }
            }
            var paths = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "config", "benefits", "manulife-billing-header-aliases.json"),
                Path.Combine(Directory.GetCurrentDirectory(), "config", "benefits", "manulife-billing-header-aliases.json")
            };
            var path = paths.FirstOrDefault(File.Exists);
            if (path is null)
            {
                return aliases;
            }

            try
            {
                var config = JsonSerializer.Deserialize<HeaderAliasConfiguration>(File.ReadAllText(path));
                foreach (var entry in config?.Aliases ?? new Dictionary<string, List<string>>())
                {
                    if (!aliases.TryGetValue(entry.Key, out var values))
                    {
                        values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        aliases[entry.Key] = values;
                    }
                    foreach (var alias in entry.Value)
                    {
                        values.Add(NormalizeHeader(alias));
                    }
                }
            }
            catch
            {
                // The canonical aliases remain available if a local configuration file is malformed.
            }

            return aliases;
        }

        private sealed class HeaderAliasConfiguration
        {
            public Dictionary<string, List<string>> Aliases { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        }
    }

    private static class BillingTableReader
    {
        public static BillingTable Read(string path, long maximumFileSizeBytes)
        {
            if (new FileInfo(path).Length > maximumFileSizeBytes)
            {
                throw new InvalidDataException("Billing file is larger than the configured intake limit.");
            }

            return Path.GetExtension(path).Equals(".csv", StringComparison.OrdinalIgnoreCase)
                ? ReadCsv(path)
                : ReadWorkbook(path);
        }

        private static BillingTable ReadWorkbook(string path)
        {
            using var workbook = new XLWorkbook(path);
            var worksheet = workbook.Worksheets.FirstOrDefault() ?? throw new InvalidDataException("The billing workbook has no worksheets.");
            var headerRow = worksheet.FirstRowUsed() ?? throw new InvalidDataException("The billing workbook has no populated header row.");
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow.RowNumber();
            var lastColumn = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;
            var table = new BillingTable();
            for (var column = 1; column <= lastColumn; column++)
            {
                table.Headers.Add(ReadCell(worksheet.Cell(headerRow.RowNumber(), column)));
            }
            for (var row = headerRow.RowNumber() + 1; row <= lastRow; row++)
            {
                var item = new BillingTableRow { RowNumber = row };
                for (var column = 1; column <= lastColumn; column++)
                {
                    item.Cells.Add(ReadCell(worksheet.Cell(row, column)));
                }
                table.Rows.Add(item);
            }
            return table;
        }

        private static BillingCell ReadCell(IXLCell cell)
        {
            var formula = cell.FormulaA1;
            return new BillingCell
            {
                Value = string.IsNullOrWhiteSpace(formula) ? cell.GetFormattedString() : formula,
                IsFormula = !string.IsNullOrWhiteSpace(formula)
            };
        }

        private static BillingTable ReadCsv(string path)
        {
            var rows = ParseCsv(File.ReadAllText(path));
            if (rows.Count == 0)
            {
                throw new InvalidDataException("The CSV billing file is empty.");
            }
            var table = new BillingTable();
            table.Headers.AddRange(rows[0].Select(value => new BillingCell { Value = value }));
            for (var index = 1; index < rows.Count; index++)
            {
                var item = new BillingTableRow { RowNumber = index + 1 };
                item.Cells.AddRange(rows[index].Select(value => new BillingCell { Value = value, IsFormula = LooksLikeFormulaInput(value) }));
                while (item.Cells.Count < table.Headers.Count)
                {
                    item.Cells.Add(new BillingCell());
                }
                table.Rows.Add(item);
            }
            return table;
        }

        private static List<List<string>> ParseCsv(string text)
        {
            var rows = new List<List<string>>();
            var row = new List<string>();
            var cell = new StringBuilder();
            var quoted = false;
            for (var index = 0; index < text.Length; index++)
            {
                var current = text[index];
                if (current == '"')
                {
                    if (quoted && index + 1 < text.Length && text[index + 1] == '"')
                    {
                        cell.Append('"');
                        index++;
                    }
                    else
                    {
                        quoted = !quoted;
                    }
                    continue;
                }
                if (current == ',' && !quoted)
                {
                    row.Add(cell.ToString());
                    cell.Clear();
                    continue;
                }
                if ((current == '\r' || current == '\n') && !quoted)
                {
                    if (current == '\r' && index + 1 < text.Length && text[index + 1] == '\n')
                    {
                        index++;
                    }
                    row.Add(cell.ToString());
                    cell.Clear();
                    rows.Add(row);
                    row = [];
                    continue;
                }
                cell.Append(current);
            }
            if (cell.Length > 0 || row.Count > 0)
            {
                row.Add(cell.ToString());
                rows.Add(row);
            }
            return rows;
        }
    }
}
