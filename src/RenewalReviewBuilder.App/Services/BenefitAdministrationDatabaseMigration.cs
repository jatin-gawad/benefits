using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class BenefitAdministrationDatabaseMigration
{
    private static readonly (string Code, string Name)[] BenefitTypes =
    [
        (CanonicalBenefitTypeCodes.Life, "Group Life Insurance"),
        (CanonicalBenefitTypeCodes.AdAndD, "Accidental Death and Dismemberment"),
        (CanonicalBenefitTypeCodes.Ltd, "Long-Term Disability"),
        (CanonicalBenefitTypeCodes.Ehc, "Extended Health Care"),
        (CanonicalBenefitTypeCodes.Travel, "Travel Insurance and Assistance"),
        (CanonicalBenefitTypeCodes.MedicalSecondOpinion, "Medical Second Opinion"),
        (CanonicalBenefitTypeCodes.VirtualHealthcare, "Virtual Healthcare"),
        (CanonicalBenefitTypeCodes.Dental, "Dental Care"),
        (CanonicalBenefitTypeCodes.OptionalLife, "Optional Life Insurance"),
        (CanonicalBenefitTypeCodes.OptionalAccidentSeriousIllness, "Optional Accident and Serious Illness"),
        (CanonicalBenefitTypeCodes.OptionalCriticalIllness, "Optional Critical Illness")
    ];

    private static readonly OrganizationSeed[] Organizations =
    [
        new("GroupHEALTH", BenefitAdministrationOrganizationTypes.ThirdPartyAdministrator),
        new("SSQ, Life Insurance Company Inc.", BenefitAdministrationOrganizationTypes.Insurer),
        new("Industrial Alliance", BenefitAdministrationOrganizationTypes.Insurer),
        new("Disability Management Institute", BenefitAdministrationOrganizationTypes.DisabilityManager),
        new("Trident Global Assistance", BenefitAdministrationOrganizationTypes.AssistanceProvider),
        new("WorldCare", BenefitAdministrationOrganizationTypes.SecondOpinionProvider),
        new("Advica Health", BenefitAdministrationOrganizationTypes.SecondOpinionProvider),
        new("Teladoc Health", BenefitAdministrationOrganizationTypes.OtherServiceProvider),
        new("Dialogue", BenefitAdministrationOrganizationTypes.OtherServiceProvider),
        new("Sparrow", BenefitAdministrationOrganizationTypes.OtherServiceProvider),
        new("TELUS Health Virtual Care", BenefitAdministrationOrganizationTypes.VirtualCareProvider),
        new("The Co-operators", BenefitAdministrationOrganizationTypes.Insurer)
    ];

    public static async Task ApplyAsync(SqliteConnection connection)
    {
        await ExecuteAsync(connection, """
            CREATE TABLE IF NOT EXISTS Organization (
                OrganizationId INTEGER PRIMARY KEY AUTOINCREMENT,
                CanonicalName TEXT NOT NULL,
                DisplayName TEXT NOT NULL,
                OrganizationTypeCode TEXT NOT NULL CHECK (OrganizationTypeCode IN (
                    'INSURER', 'CARRIER', 'THIRD_PARTY_ADMINISTRATOR', 'DISABILITY_MANAGER',
                    'ASSISTANCE_PROVIDER', 'VIRTUAL_CARE_PROVIDER', 'SECOND_OPINION_PROVIDER',
                    'OTHER_SERVICE_PROVIDER')),
                NormalizedName TEXT,
                IsActive INTEGER NOT NULL DEFAULT 1,
                CreatedAtUtc TEXT NOT NULL,
                UpdatedAtUtc TEXT NOT NULL
            );
            """);

        await ExecuteAsync(connection, """
            CREATE UNIQUE INDEX IF NOT EXISTS UX_Organization_NormalizedName
            ON Organization(NormalizedName COLLATE NOCASE)
            WHERE NormalizedName IS NOT NULL AND TRIM(NormalizedName) <> '';

            CREATE INDEX IF NOT EXISTS IX_Organization_Type
            ON Organization(OrganizationTypeCode, IsActive);
            """);

        await ExecuteAsync(connection, """
            CREATE TABLE IF NOT EXISTS ClientBenefitAdministration (
                ClientBenefitAdministrationId INTEGER PRIMARY KEY AUTOINCREMENT,
                ClientId INTEGER NOT NULL,
                PlanDesignVersionId INTEGER,
                BenefitTypeId INTEGER NOT NULL,
                BenefitDisplayName TEXT,
                ProductName TEXT,
                PolicyNumber TEXT,
                InsurerOrganizationId INTEGER,
                InsurerSourceDisplayName TEXT,
                IsOptional INTEGER NOT NULL DEFAULT 0,
                EffectiveDate TEXT,
                TerminationDate TEXT,
                SourceDocumentId INTEGER,
                SourcePageNumber INTEGER,
                SourceText TEXT,
                ExtractionConfidence REAL,
                ReviewStatusCode TEXT NOT NULL DEFAULT 'NEEDS_REVIEW'
                    CHECK (ReviewStatusCode IN ('NEEDS_REVIEW', 'CONFIRMED', 'REJECTED')),
                IsActive INTEGER NOT NULL DEFAULT 1,
                CopiedFromClientBenefitAdministrationId INTEGER,
                CreatedAtUtc TEXT NOT NULL,
                UpdatedAtUtc TEXT NOT NULL,
                FOREIGN KEY (ClientId) REFERENCES Clients(client_id),
                FOREIGN KEY (BenefitTypeId) REFERENCES Benefit(benefit_id),
                FOREIGN KEY (InsurerOrganizationId) REFERENCES Organization(OrganizationId),
                FOREIGN KEY (SourceDocumentId) REFERENCES Documents(document_id),
                FOREIGN KEY (CopiedFromClientBenefitAdministrationId)
                    REFERENCES ClientBenefitAdministration(ClientBenefitAdministrationId)
            );
            """);

        await ExecuteAsync(connection, """
            CREATE TABLE IF NOT EXISTS BenefitServiceProviderAssignment (
                BenefitServiceProviderAssignmentId INTEGER PRIMARY KEY AUTOINCREMENT,
                ClientBenefitAdministrationId INTEGER NOT NULL,
                ProviderOrganizationId INTEGER NOT NULL,
                ProviderRoleCode TEXT NOT NULL CHECK (ProviderRoleCode IN (
                    'BENEFIT_ADMINISTRATOR', 'CLAIMS_ADMINISTRATOR', 'DISABILITY_MANAGEMENT',
                    'TRAVEL_ASSISTANCE', 'MEDICAL_SECOND_OPINION', 'VIRTUAL_CARE',
                    'OPTIONAL_BENEFIT_ADMINISTRATOR', 'OTHER')),
                SourceDisplayName TEXT,
                IsPrimary INTEGER NOT NULL DEFAULT 1,
                EffectiveDate TEXT,
                TerminationDate TEXT,
                SourceDocumentId INTEGER,
                SourcePageNumber INTEGER,
                ReviewStatusCode TEXT NOT NULL DEFAULT 'NEEDS_REVIEW'
                    CHECK (ReviewStatusCode IN ('NEEDS_REVIEW', 'CONFIRMED', 'REJECTED')),
                IsActive INTEGER NOT NULL DEFAULT 1,
                CreatedAtUtc TEXT NOT NULL,
                UpdatedAtUtc TEXT NOT NULL,
                FOREIGN KEY (ClientBenefitAdministrationId)
                    REFERENCES ClientBenefitAdministration(ClientBenefitAdministrationId),
                FOREIGN KEY (ProviderOrganizationId) REFERENCES Organization(OrganizationId),
                FOREIGN KEY (SourceDocumentId) REFERENCES Documents(document_id)
            );
            """);

        await ExecuteAsync(connection, """
            CREATE INDEX IF NOT EXISTS IX_ClientBenefitAdministration_Client
            ON ClientBenefitAdministration(ClientId);
            CREATE INDEX IF NOT EXISTS IX_ClientBenefitAdministration_PlanDesignVersion
            ON ClientBenefitAdministration(PlanDesignVersionId);
            CREATE INDEX IF NOT EXISTS IX_ClientBenefitAdministration_BenefitType
            ON ClientBenefitAdministration(BenefitTypeId);
            CREATE INDEX IF NOT EXISTS IX_ClientBenefitAdministration_PolicyNumber
            ON ClientBenefitAdministration(PolicyNumber COLLATE NOCASE);
            CREATE INDEX IF NOT EXISTS IX_ClientBenefitAdministration_Insurer
            ON ClientBenefitAdministration(InsurerOrganizationId);
            CREATE INDEX IF NOT EXISTS IX_BenefitServiceProviderAssignment_Administration
            ON BenefitServiceProviderAssignment(ClientBenefitAdministrationId);
            CREATE INDEX IF NOT EXISTS IX_BenefitServiceProviderAssignment_Provider
            ON BenefitServiceProviderAssignment(ProviderOrganizationId);

            CREATE UNIQUE INDEX IF NOT EXISTS UX_ClientBenefitAdministration_ActiveIdentity
            ON ClientBenefitAdministration(
                ClientId,
                COALESCE(PlanDesignVersionId, 0),
                BenefitTypeId,
                LOWER(TRIM(COALESCE(ProductName, ''))),
                COALESCE(EffectiveDate, '')
            )
            WHERE IsActive = 1;

            CREATE UNIQUE INDEX IF NOT EXISTS UX_BenefitServiceProviderAssignment_ActiveIdentity
            ON BenefitServiceProviderAssignment(
                ClientBenefitAdministrationId,
                ProviderOrganizationId,
                ProviderRoleCode,
                COALESCE(EffectiveDate, '')
            )
            WHERE IsActive = 1;
            """);

        foreach (var organization in Organizations)
        {
            await SeedOrganizationAsync(connection, organization);
        }

        foreach (var benefitType in BenefitTypes)
        {
            await SeedBenefitTypeAsync(connection, benefitType.Code, benefitType.Name);
        }

        await EnsureBenefitTypeIndexAsync(connection);
    }

    private static async Task SeedOrganizationAsync(SqliteConnection connection, OrganizationSeed organization)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = """
            INSERT INTO Organization (
                CanonicalName, DisplayName, OrganizationTypeCode, NormalizedName,
                IsActive, CreatedAtUtc, UpdatedAtUtc)
            SELECT $canonical_name, $display_name, $type_code, $normalized_name,
                   1, $now, $now
            WHERE NOT EXISTS (
                SELECT 1 FROM Organization
                WHERE NormalizedName = $normalized_name COLLATE NOCASE
            );
            """;
        command.Parameters.AddWithValue("$canonical_name", organization.Name);
        command.Parameters.AddWithValue("$display_name", organization.Name);
        command.Parameters.AddWithValue("$type_code", organization.TypeCode);
        command.Parameters.AddWithValue("$normalized_name", NormalizeName(organization.Name));
        command.Parameters.AddWithValue("$now", DateTime.UtcNow.ToString("O"));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task SeedBenefitTypeAsync(SqliteConnection connection, string code, string name)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = """
            INSERT INTO Benefit (
                benefit_category, benefit_line, normalized_name,
                source_type, imported_at, extraction_method)
            SELECT 'Canonical Benefit Type', $benefit_line, $normalized_name,
                   'System', $now, 'BenefitAdministrationMigration'
            WHERE NOT EXISTS (
                SELECT 1 FROM Benefit WHERE benefit_line = $benefit_line COLLATE NOCASE
            );
            """;
        command.Parameters.AddWithValue("$benefit_line", code);
        command.Parameters.AddWithValue("$normalized_name", name);
        command.Parameters.AddWithValue("$now", DateTime.UtcNow.ToString("O"));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task EnsureBenefitTypeIndexAsync(SqliteConnection connection)
    {
        await using var duplicateCommand = connection.CreateCommand();
        duplicateCommand.CommandText = """
            SELECT COUNT(*)
            FROM (
                SELECT LOWER(TRIM(benefit_line))
                FROM Benefit
                WHERE TRIM(COALESCE(benefit_line, '')) <> ''
                GROUP BY LOWER(TRIM(benefit_line))
                HAVING COUNT(*) > 1
            );
            """;
        var duplicateCount = Convert.ToInt32(await duplicateCommand.ExecuteScalarAsync());
        if (duplicateCount == 0)
        {
            await ExecuteAsync(connection, """
                CREATE UNIQUE INDEX IF NOT EXISTS UX_Benefit_CanonicalLine
                ON Benefit(benefit_line COLLATE NOCASE)
                WHERE TRIM(COALESCE(benefit_line, '')) <> '';
                """);
        }
        else
        {
            AppLogger.Info("Benefit canonical-line uniqueness was not added because legacy duplicate benefit rows require review.");
        }
    }

    private static async Task ExecuteAsync(SqliteConnection connection, string sql)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        await command.ExecuteNonQueryAsync();
    }

    private static string NormalizeName(string value)
    {
        return new string(value.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
    }

    private sealed record OrganizationSeed(string Name, string TypeCode);
}
