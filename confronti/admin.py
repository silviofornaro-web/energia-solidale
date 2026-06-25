from django.contrib import admin

from .models import ComparisonReport, CustomerArchiveFolder, InviteCode


@admin.register(InviteCode)
class InviteCodeAdmin(admin.ModelAdmin):
    list_display = ("code", "label", "is_active", "used_at", "used_by", "created_at")
    list_filter = ("is_active", "created_at", "used_at")
    search_fields = ("code", "label", "note", "used_by__username", "used_by__email")
    readonly_fields = ("created_at", "used_at", "used_by")


class ComparisonReportInline(admin.TabularInline):
    model = ComparisonReport
    extra = 0
    fields = ("title", "providers_label", "commodity", "comparison_datetime", "report_file", "created_at")
    readonly_fields = ("created_at",)


@admin.register(CustomerArchiveFolder)
class CustomerArchiveFolderAdmin(admin.ModelAdmin):
    list_display = ("customer_name", "folder_name", "customer_email", "customer_phone", "created_at", "updated_at")
    list_filter = ("created_at", "updated_at")
    search_fields = ("customer_name", "customer_email", "customer_phone", "folder_name", "notes")
    readonly_fields = ("folder_name", "created_at", "updated_at", "created_by")
    inlines = [ComparisonReportInline]


@admin.register(ComparisonReport)
class ComparisonReportAdmin(admin.ModelAdmin):
    list_display = (
        "title",
        "folder",
        "providers_label",
        "commodity",
        "comparison_datetime",
        "created_at",
    )
    list_filter = ("commodity", "created_at", "comparison_datetime")
    search_fields = ("title", "folder__customer_name", "folder__folder_name", "providers_label", "notes")
    readonly_fields = ("created_at", "updated_at", "created_by")
