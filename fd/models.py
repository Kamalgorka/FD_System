from django.db import models, transaction


# ------------------------------
# Bank Master
# ------------------------------
class BankMaster(models.Model):
    bank_name = models.CharField(max_length=200)
    fd_made_bank = models.CharField(max_length=200, blank=True, null=True)
    category = models.CharField(max_length=200, blank=True, null=True)
    sort_order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ["sort_order", "bank_name"]
        unique_together = ("bank_name",)
        verbose_name = "Bank Master"
        verbose_name_plural = "Bank Masters"

    def __str__(self):
        return f"{self.bank_name}{f' ({self.category})' if self.category else ''}"


# ------------------------------
# Trustee Master
# ------------------------------
class TrusteeMaster(models.Model):
    name = models.CharField(max_length=200, unique=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ["name"]
        verbose_name = "Trustee Master"
        verbose_name_plural = "Trustee Masters"

    def __str__(self):
        return self.name


# ------------------------------
# Financial Year Helper
# ------------------------------
def get_financial_year(dt):
    y = dt.year
    if dt.month >= 4:
        start = y
        end = y + 1
    else:
        start = y - 1
        end = y
    return f"{str(start)[-2:]}-{str(end)[-2:]}"


# ------------------------------
# FD Number Sequence
# ------------------------------
class FDNumberSequence(models.Model):
    fy = models.CharField(max_length=5)
    fd_type_code = models.CharField(max_length=3)
    last_no = models.PositiveIntegerField(default=0)

    class Meta:
        unique_together = ("fy", "fd_type_code")

    def __str__(self):
        return f"{self.fy}-{self.fd_type_code}-{self.last_no}"


# ------------------------------
# FD Entry
# ------------------------------
class FDEntry(models.Model):
    system_fd_no = models.PositiveIntegerField(unique=True, null=True, blank=True)
    fd_auto_code = models.CharField(max_length=30, unique=True, null=True, blank=True)

    fd_number_receipt = models.CharField(max_length=100, blank=True, null=True)

    FD_TYPE_CHOICES = (
        ("With lien", "With lien"),
        ("Without lien", "Without lien"),
    )
    fd_type = models.CharField(max_length=50, choices=FD_TYPE_CHOICES, blank=True, null=True)

    FD_NATURE_CHOICES = (
        ("Callable", "Callable"),
        ("Non Callable", "Non Callable"),
    )
    fd_nature = models.CharField(max_length=50, choices=FD_NATURE_CHOICES, blank=True, null=True)

    term_loan_number = models.CharField(max_length=100, blank=True, null=True)
    term_loan_with = models.CharField(max_length=200, blank=True, null=True)

    od_status = models.BooleanField(default=False)
    margin_of_od = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    rate_of_od = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    amount_od = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)

    bank_master = models.ForeignKey(BankMaster, on_delete=models.PROTECT, null=True, blank=True)
    trustee = models.ForeignKey(TrusteeMaster, null=True, blank=True, on_delete=models.PROTECT)

    start_date = models.DateField()
    maturity_date = models.DateField()

    days = models.IntegerField(default=0)
    days_bucket = models.CharField(max_length=100, blank=True, null=True)

    roi = models.DecimalField(max_digits=6, decimal_places=2)
    percentage_bucket = models.CharField(max_length=50, blank=True, null=True)

    fd_amount = models.DecimalField(max_digits=15, decimal_places=2)
    maturity_amount = models.DecimalField(max_digits=15, decimal_places=2)
    interest_actually_due = models.DecimalField(max_digits=15, decimal_places=2)

    tds_percent = models.DecimalField(max_digits=6, decimal_places=2)
    tds_amount = models.DecimalField(max_digits=15, decimal_places=2, default=0)

    amt_expected = models.DecimalField(max_digits=15, decimal_places=2, default=0)

    remarks = models.TextField(blank=True, null=True)
    attachment = models.FileField(upload_to="fd_attachments/", blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "FD Entry"
        verbose_name_plural = "FD Entries"

    def __str__(self):
        bank = self.bank_master.bank_name if self.bank_master else ""
        return f"{self.fd_auto_code or ''} - {bank}"

    @property
    def system_fd_code(self):
        n = self.system_fd_no or 0
        return f"MML-{n:03d}"

    def save(self, *args, **kwargs):
        if not self.system_fd_no:
            with transaction.atomic():
                last = FDEntry.objects.select_for_update().order_by("-system_fd_no").first()
                self.system_fd_no = (last.system_fd_no if last and last.system_fd_no else 0) + 1
        super().save(*args, **kwargs)