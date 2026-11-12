namespace RenewalReviewBuilder.App.Models;

/// <summary>
/// Master bonus program configuration for a carrier and program year.
/// </summary>
public sealed class BonusProgramRecord
{
    /// <summary>
    /// Gets or sets the SQLite primary key for the bonus program.
    /// </summary>
    public long BonusProgramId { get; set; }

    /// <summary>
    /// Gets or sets the linked Carrier master table primary key.
    /// </summary>
    public long CarrierId { get; set; }

    /// <summary>
    /// Gets or sets the carrier display name from the Carrier table when records are queried.
    /// </summary>
    public string CarrierName { get; set; } = "";

    /// <summary>
    /// Gets or sets the bonus program name.
    /// </summary>
    public string ProgramName { get; set; } = "";

    /// <summary>
    /// Gets or sets the bonus program year.
    /// </summary>
    public int ProgramYear { get; set; }

    /// <summary>
    /// Gets or sets the first qualification date for this program.
    /// </summary>
    public DateOnly QualificationStartDate { get; set; }

    /// <summary>
    /// Gets or sets the last qualification date for this program.
    /// </summary>
    public DateOnly QualificationEndDate { get; set; }

    /// <summary>
    /// Gets or sets the optional payment frequency.
    /// </summary>
    public string? PaymentFrequency { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes a sales bonus.
    /// </summary>
    public bool HasSalesBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes a persistency bonus.
    /// </summary>
    public bool HasPersistencyBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes a growth bonus.
    /// </summary>
    public bool HasGrowthBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes a retention bonus.
    /// </summary>
    public bool HasRetentionBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes an elite bonus.
    /// </summary>
    public bool HasEliteBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes a sales enhancement.
    /// </summary>
    public bool HasSalesEnhancement { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether rate guarantees can affect this program.
    /// </summary>
    public bool HasRateGuaranteeImpact { get; set; }

    /// <summary>
    /// Gets or sets the source document name or path used to define this program.
    /// </summary>
    public string? SourceDocument { get; set; }

    /// <summary>
    /// Gets or sets optional implementation notes.
    /// </summary>
    public string? Notes { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this program is active.
    /// </summary>
    public bool IsActive { get; set; } = true;

    /// <summary>
    /// Gets or sets the date and time the program row was created.
    /// </summary>
    public DateTime CreatedDate { get; set; }

    /// <summary>
    /// Gets or sets the date and time the program row was last modified.
    /// </summary>
    public DateTime? ModifiedDate { get; set; }

    /// <summary>
    /// Creates a simple transfer object from this entity.
    /// </summary>
    public BonusProgramDto ToDto()
    {
        return new BonusProgramDto
        {
            BonusProgramId = BonusProgramId,
            CarrierId = CarrierId,
            CarrierName = CarrierName,
            ProgramName = ProgramName,
            ProgramYear = ProgramYear,
            QualificationStartDate = QualificationStartDate,
            QualificationEndDate = QualificationEndDate,
            PaymentFrequency = PaymentFrequency,
            HasSalesBonus = HasSalesBonus,
            HasPersistencyBonus = HasPersistencyBonus,
            HasGrowthBonus = HasGrowthBonus,
            HasRetentionBonus = HasRetentionBonus,
            HasEliteBonus = HasEliteBonus,
            HasSalesEnhancement = HasSalesEnhancement,
            HasRateGuaranteeImpact = HasRateGuaranteeImpact,
            SourceDocument = SourceDocument,
            Notes = Notes,
            IsActive = IsActive,
            CreatedDate = CreatedDate,
            ModifiedDate = ModifiedDate
        };
    }
}

/// <summary>
/// Bonus program data transfer object for service and future UI/API boundaries.
/// </summary>
public sealed class BonusProgramDto
{
    /// <summary>
    /// Gets or sets the SQLite primary key for the bonus program.
    /// </summary>
    public long BonusProgramId { get; set; }

    /// <summary>
    /// Gets or sets the linked Carrier master table primary key.
    /// </summary>
    public long CarrierId { get; set; }

    /// <summary>
    /// Gets or sets the carrier display name.
    /// </summary>
    public string CarrierName { get; set; } = "";

    /// <summary>
    /// Gets or sets the bonus program name.
    /// </summary>
    public string ProgramName { get; set; } = "";

    /// <summary>
    /// Gets or sets the bonus program year.
    /// </summary>
    public int ProgramYear { get; set; }

    /// <summary>
    /// Gets or sets the first qualification date for this program.
    /// </summary>
    public DateOnly QualificationStartDate { get; set; }

    /// <summary>
    /// Gets or sets the last qualification date for this program.
    /// </summary>
    public DateOnly QualificationEndDate { get; set; }

    /// <summary>
    /// Gets or sets the optional payment frequency.
    /// </summary>
    public string? PaymentFrequency { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes a sales bonus.
    /// </summary>
    public bool HasSalesBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes a persistency bonus.
    /// </summary>
    public bool HasPersistencyBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes a growth bonus.
    /// </summary>
    public bool HasGrowthBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes a retention bonus.
    /// </summary>
    public bool HasRetentionBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes an elite bonus.
    /// </summary>
    public bool HasEliteBonus { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the program includes a sales enhancement.
    /// </summary>
    public bool HasSalesEnhancement { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether rate guarantees can affect this program.
    /// </summary>
    public bool HasRateGuaranteeImpact { get; set; }

    /// <summary>
    /// Gets or sets the source document name or path used to define this program.
    /// </summary>
    public string? SourceDocument { get; set; }

    /// <summary>
    /// Gets or sets optional implementation notes.
    /// </summary>
    public string? Notes { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this program is active.
    /// </summary>
    public bool IsActive { get; set; }

    /// <summary>
    /// Gets or sets the date and time the program row was created.
    /// </summary>
    public DateTime CreatedDate { get; set; }

    /// <summary>
    /// Gets or sets the date and time the program row was last modified.
    /// </summary>
    public DateTime? ModifiedDate { get; set; }
}
