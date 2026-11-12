using Microsoft.Data.Sqlite;

namespace RenewalReviewBuilder.App.Services;

public static class PlanBenefitRuleDatabaseMigration
{
    public static async Task ApplyAsync(SqliteConnection connection)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = """
            CREATE TABLE IF NOT EXISTS PlanBenefitRule (
                PlanBenefitRuleId INTEGER PRIMARY KEY AUTOINCREMENT,
                PlanDesignFeatureId INTEGER NOT NULL,
                ClientId INTEGER NOT NULL,
                SourceDocumentId INTEGER,
                RuleType TEXT NOT NULL CHECK (RuleType IN ('TieredReimbursement')),
                Tier1ReimbursementPercent REAL,
                Tier1ThresholdAmount REAL,
                Tier2ReimbursementPercent REAL,
                ThresholdBasis TEXT CHECK (ThresholdBasis IS NULL OR ThresholdBasis IN ('Person', 'Certificate')),
                ResetPeriod TEXT CHECK (ResetPeriod IS NULL OR ResetPeriod IN ('CalendarYear', 'PolicyYear')),
                CurrencyCode TEXT,
                SourceText TEXT,
                SourceLocation TEXT,
                NormalizationStatus TEXT NOT NULL DEFAULT 'NeedsReview'
                    CHECK (NormalizationStatus IN ('Normalized', 'NeedsReview')),
                ConfidenceScore REAL,
                DisplayOrder INTEGER NOT NULL DEFAULT 10,
                IsActive INTEGER NOT NULL DEFAULT 1,
                EffectiveDate TEXT,
                TerminationDate TEXT,
                CreatedAtUtc TEXT NOT NULL,
                UpdatedAtUtc TEXT NOT NULL,
                FOREIGN KEY (PlanDesignFeatureId) REFERENCES plan_design_features(plan_design_feature_id) ON DELETE CASCADE,
                FOREIGN KEY (ClientId) REFERENCES Clients(client_id),
                FOREIGN KEY (SourceDocumentId) REFERENCES Documents(document_id)
            );

            CREATE INDEX IF NOT EXISTS IX_PlanBenefitRule_PlanDesignFeature
            ON PlanBenefitRule(PlanDesignFeatureId, IsActive, DisplayOrder);

            CREATE INDEX IF NOT EXISTS IX_PlanBenefitRule_Client
            ON PlanBenefitRule(ClientId, RuleType, IsActive);

            CREATE INDEX IF NOT EXISTS IX_PlanBenefitRule_SourceDocument
            ON PlanBenefitRule(SourceDocumentId, RuleType);

            CREATE UNIQUE INDEX IF NOT EXISTS UX_PlanBenefitRule_ActiveFeatureType
            ON PlanBenefitRule(PlanDesignFeatureId, RuleType)
            WHERE IsActive = 1;
            """;
        await command.ExecuteNonQueryAsync();
    }
}
