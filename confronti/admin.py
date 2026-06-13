from django.contrib import admin

from .models import InviteCode


@admin.register(InviteCode)
class InviteCodeAdmin(admin.ModelAdmin):
    list_display = ("code", "label", "is_active", "used_at", "used_by", "created_at")
    list_filter = ("is_active", "created_at", "used_at")
    search_fields = ("code", "label", "used_by__username", "used_by__email")
    readonly_fields = ("created_at", "used_at", "used_by")
