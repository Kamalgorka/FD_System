from django.contrib import admin
from django.urls import path, include
from django.shortcuts import redirect   # 👈 ADD THIS

from django.conf import settings
from django.conf.urls.static import static

# 👇 ADD THIS FUNCTION
def home(request):
    return redirect('/fd/dashboard/')

urlpatterns = [
    path('', home),   # 👈 ADD THIS LINE
    path("admin/", admin.site.urls),
    path("fd/", include("fd.urls")),
]

# Serve uploaded files in development only
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)