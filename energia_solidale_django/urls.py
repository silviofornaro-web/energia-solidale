from django.contrib import admin
from django.contrib.auth import views as auth_views
from django.urls import include, path

from confronti.auth_views import ResilientLoginView
from confronti.diagnostics import db_health


urlpatterns = [
    path("admin/", admin.site.urls),
    path("__health/db/", db_health, name="db_health"),
    path("accounts/login/", ResilientLoginView.as_view(template_name="registration/login.html"), name="login"),
    path("accounts/logout/", auth_views.LogoutView.as_view(), name="logout"),
    path("", include("confronti.urls")),
]
