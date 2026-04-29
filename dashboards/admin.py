from django.contrib import admin
from django.shortcuts import redirect

from .models import DailyDashboardNav, DailyLiquidityInput


class DailyDashboardNavAdmin(admin.ModelAdmin):
    def changelist_view(self, request, extra_context=None):
        return redirect("/fd/dashboard/")

    def has_add_permission(self, request):
        return False

    def has_delete_permission(self, request, obj=None):
        return False


class DailyLiquidityInputAdmin(admin.ModelAdmin):
    list_display = ("as_on_date", "funds_liquid_option", "funds_overnight_option", "funds_ncds")
    search_fields = ("as_on_date",)
    list_filter = ("as_on_date",)
    ordering = ("-as_on_date",)


admin.site.register(DailyDashboardNav, DailyDashboardNavAdmin)
admin.site.register(DailyLiquidityInput, DailyLiquidityInputAdmin)
