from django.db import models


class DailyDashboardNav(models.Model):
    """
    Dummy model only to create a sidebar link in admin.
    Clicking it redirects to /fd/dashboard/
    """

    class Meta:
        verbose_name = "Daily Dashboard"
        verbose_name_plural = "Daily Dashboard"

    def __str__(self):
        return "Daily Dashboard"


class DailyLiquidityInput(models.Model):
    """
    Stores manual liquidity inputs date-wise.
    """

    as_on_date = models.DateField(unique=True)

    funds_liquid_option = models.DecimalField(
        max_digits=14,
        decimal_places=2,
        default=0
    )

    funds_overnight_option = models.DecimalField(
        max_digits=14,
        decimal_places=2,
        default=0
    )

    funds_ncds = models.DecimalField(
        max_digits=14,
        decimal_places=2,
        default=0
    )

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Liquidity Input (Daily)"
        verbose_name_plural = "Liquidity Inputs (Daily)"
        ordering = ["-as_on_date"]

    def __str__(self):
        return f"Liquidity {self.as_on_date}"
