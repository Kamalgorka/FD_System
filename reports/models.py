from django.db import models
from fd.models import FDEntry


class ReportCenter(FDEntry):
    """
    Only for Reports screen (real-time). Uses same FDEntry table.
    """
    class Meta:
        proxy = True
        verbose_name = "Reports"
        verbose_name_plural = "Reports"


# ---------------------------------------------------------
# NEW: Interest Income Quarter-wise Upload Models
# ---------------------------------------------------------

class InterestIncomeUpload(models.Model):
    Q_CHOICES = [
        ("Q1", "Q1 (Apr–Jun)"),
        ("Q2", "Q2 (Jul–Sep)"),
        ("Q3", "Q3 (Oct–Dec)"),
        ("Q4", "Q4 (Jan–Mar)"),
    ]

    quarter = models.CharField(max_length=2, choices=Q_CHOICES)
    as_on_date = models.DateField()
    upload_file = models.FileField(upload_to="interest_income_uploads/%Y/%m/")
    uploaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-uploaded_at"]
        verbose_name = "Interest Income Data"
        verbose_name_plural = "Interest Income Data"

    def __str__(self):
        return f"Interest Income {self.quarter} | As on {self.as_on_date}"


class InterestIncomeLine(models.Model):
    upload = models.ForeignKey(
        InterestIncomeUpload, on_delete=models.CASCADE, related_name="lines"
    )

    # Common key (must match FDEntry.fd_number_receipt)
    fd_number_receipt = models.CharField(max_length=150, db_index=True)
    is_matched_fd = models.BooleanField(default=False)

    # Interest Income during the year (7 cols)
    opening_interest_accrued = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    additions_during_year = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    unbooked_interest_income_on_maturity = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    total_income_on_fd_at_maturity_actual = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    interest_income_self_calculated = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    difference_t_u = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    term_of_maturity = models.CharField(max_length=120, null=True, blank=True)

    # Recoverable details (4 cols)
    recoverable_opening = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    recoverable_additions = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    recoverable_deletion = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    recoverable_closing = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)

    # TDS Deducted (4 cols)
    tds_opening = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    tds_additions_during_year = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    tds_additions_on_unbooked_tds = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    tds_closing = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)

    # Extra (2 cols)
    interest_as_per_calculation = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    basis_of_calculation = models.TextField(null=True, blank=True)

    class Meta:
        unique_together = ("upload", "fd_number_receipt")
        ordering = ["fd_number_receipt"]
        verbose_name = "Interest Income Line"
        verbose_name_plural = "Interest Income Lines"

    def __str__(self):
        return f"{self.fd_number_receipt} ({self.upload.quarter})"
