namespace RenewalReviewBuilder.App.Models;

/// <summary>
/// Master insurance carrier reference used by Commission Intelligence and other carrier-aware modules.
/// </summary>
public sealed class CarrierRecord
{
    /// <summary>
    /// Gets or sets the SQLite primary key for the carrier.
    /// </summary>
    public long CarrierId { get; set; }

    /// <summary>
    /// Gets or sets the stable short code used to reference the carrier.
    /// </summary>
    public string CarrierCode { get; set; } = "";

    /// <summary>
    /// Gets or sets the display name for the insurance carrier.
    /// </summary>
    public string CarrierName { get; set; } = "";

    /// <summary>
    /// Gets or sets the commission statement parser key for imported commission files.
    /// </summary>
    public string CommissionStatementFormat { get; set; } = "";

    /// <summary>
    /// Gets or sets the optional bonus program version used for carrier-specific bonus rules.
    /// </summary>
    public string? BonusProgramVersion { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this carrier has rate guarantee rules.
    /// </summary>
    public bool HasRateGuarantees { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this carrier supports persistency bonus tracking.
    /// </summary>
    public bool HasPersistencyBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this carrier supports sales bonus tracking.
    /// </summary>
    public bool HasSalesBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this carrier supports growth bonus tracking.
    /// </summary>
    public bool HasGrowthBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether monthly billing imports are available for this carrier.
    /// </summary>
    public bool SupportsBillingImport { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether claims experience imports are available for this carrier.
    /// </summary>
    public bool SupportsClaimsImport { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this carrier is active for future workflows.
    /// </summary>
    public bool IsActive { get; set; } = true;

    /// <summary>
    /// Gets or sets the date and time the carrier row was created.
    /// </summary>
    public DateTime CreatedDate { get; set; }

    /// <summary>
    /// Gets or sets the date and time the carrier row was last modified.
    /// </summary>
    public DateTime? ModifiedDate { get; set; }

    /// <summary>
    /// Creates a simple transfer object from this entity.
    /// </summary>
    public CarrierDto ToDto()
    {
        return new CarrierDto
        {
            CarrierId = CarrierId,
            CarrierCode = CarrierCode,
            CarrierName = CarrierName,
            CommissionStatementFormat = CommissionStatementFormat,
            BonusProgramVersion = BonusProgramVersion,
            HasRateGuarantees = HasRateGuarantees,
            HasPersistencyBonus = HasPersistencyBonus,
            HasSalesBonus = HasSalesBonus,
            HasGrowthBonus = HasGrowthBonus,
            SupportsBillingImport = SupportsBillingImport,
            SupportsClaimsImport = SupportsClaimsImport,
            IsActive = IsActive,
            CreatedDate = CreatedDate,
            ModifiedDate = ModifiedDate
        };
    }
}

/// <summary>
/// Carrier data transfer object for service and future UI/API boundaries.
/// </summary>
public sealed class CarrierDto
{
    /// <summary>
    /// Gets or sets the SQLite primary key for the carrier.
    /// </summary>
    public long CarrierId { get; set; }

    /// <summary>
    /// Gets or sets the stable short code used to reference the carrier.
    /// </summary>
    public string CarrierCode { get; set; } = "";

    /// <summary>
    /// Gets or sets the display name for the insurance carrier.
    /// </summary>
    public string CarrierName { get; set; } = "";

    /// <summary>
    /// Gets or sets the commission statement parser key for imported commission files.
    /// </summary>
    public string CommissionStatementFormat { get; set; } = "";

    /// <summary>
    /// Gets or sets the optional bonus program version used for carrier-specific bonus rules.
    /// </summary>
    public string? BonusProgramVersion { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this carrier has rate guarantee rules.
    /// </summary>
    public bool HasRateGuarantees { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this carrier supports persistency bonus tracking.
    /// </summary>
    public bool HasPersistencyBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this carrier supports sales bonus tracking.
    /// </summary>
    public bool HasSalesBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this carrier supports growth bonus tracking.
    /// </summary>
    public bool HasGrowthBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether monthly billing imports are available for this carrier.
    /// </summary>
    public bool SupportsBillingImport { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether claims experience imports are available for this carrier.
    /// </summary>
    public bool SupportsClaimsImport { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this carrier is active for future workflows.
    /// </summary>
    public bool IsActive { get; set; }

    /// <summary>
    /// Gets or sets the date and time the carrier row was created.
    /// </summary>
    public DateTime CreatedDate { get; set; }

    /// <summary>
    /// Gets or sets the date and time the carrier row was last modified.
    /// </summary>
    public DateTime? ModifiedDate { get; set; }
}
