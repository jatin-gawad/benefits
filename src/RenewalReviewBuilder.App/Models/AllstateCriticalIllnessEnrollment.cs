namespace RenewalReviewBuilder.App.Models;

/// <summary>Enrollment counts needed to price Allstate's per-employee Critical Illness tiers.</summary>
public sealed class AllstateCriticalIllnessEnrollment
{
    public int EmployeeOnlyCount { get; set; }
    public int FamilyCount { get; set; }

    public bool HasEntries => EmployeeOnlyCount > 0 || FamilyCount > 0;
    public int TotalEnrolled => EmployeeOnlyCount + FamilyCount;

    public AllstateCriticalIllnessEnrollment Clone() => new()
    {
        EmployeeOnlyCount = EmployeeOnlyCount,
        FamilyCount = FamilyCount
    };
}
