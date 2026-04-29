from django.urls import path
from . import views
from .views import download_dashboard_test
from .views import dashboard_view
urlpatterns = [
path("create/", views.fd_create, name="fd_create"),
path("bankmaster-options/", views.bankmaster_options, name="bankmaster_options"),
path("trustee-options/", views.trustee_options, name="trustee_options"),
path("fd-auto-preview/", views.fd_auto_preview, name="fd_auto_preview"),
path("dashboard/test-download/", download_dashboard_test, name="dashboard_test_download"),
path("dashboard/", dashboard_view, name="fd_dashboard"),
path("dashboard/download/pdf/", views.download_dashboard_pdf, name="download_dashboard_pdf"),
path("dashboard/download/excel/", views.download_dashboard_excel, name="download_dashboard_excel"),
path("liquidity-dashboard/excel/", views.download_dashboard_excel, name="download_dashboard_excel"),
path("liquidity-dashboard/pdf/", views.download_dashboard_pdf, name="download_dashboard_pdf"),
path("trustee-options/", views.trustee_options, name="trustee_options"),

]
