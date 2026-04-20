from django.contrib import admin
from .models import (Customer, LoadTest, Inventory, RepairInspection,
                     Certificate, TestEquipmentRow, LoadCellCertRow,
                     LoadCellSavedCert, PhotoReference)


@admin.register(Customer)
class CustomerAdmin(admin.ModelAdmin):
    list_display  = ('name', 'address')
    search_fields = ('name', 'address')


@admin.register(LoadTest)
class LoadTestAdmin(admin.ModelAdmin):
    list_display  = ('customer_name', 'company', 'description', 'serial_number',
                     'certificate_number', 'mechanic', 'date_inspection', 'date_due', 'created_at')
    search_fields = ('customer_name', 'company', 'description', 'serial_number',
                     'certificate_number', 'mechanic')
    list_filter   = ('date_inspection',)


@admin.register(Inventory)
class InventoryAdmin(admin.ModelAdmin):
    list_display  = (
        'customer_name',
        'description',
        'part_number',   # ✔ existing field
        'brand',         # ✔ existing field
        'location',
        'quantity',
        'created_at'
    )
    search_fields = (
        'customer_name',
        'description',
        'part_number',
        'brand',
        'location'
    )


@admin.register(RepairInspection)
class RepairInspectionAdmin(admin.ModelAdmin):
    list_display  = ('customer_name', 'record_type', 'company', 'description',
                     'serial_number', 'report_number', 'mechanic', 'date', 'created_at')
    search_fields = ('customer_name', 'company', 'description', 'serial_number',
                     'report_number', 'mechanic')
    list_filter   = ('record_type', 'date')


class TestEquipmentRowInline(admin.TabularInline):
    model  = TestEquipmentRow
    extra  = 1
    fields = ('description', 'capacity', 'model_no', 'serial_no', 'part_no', 'expiry_date')


class LoadCellCertRowInline(admin.TabularInline):
    model  = LoadCellCertRow
    extra  = 1
    fields = ('cert_number',)


@admin.register(Certificate)
class CertificateAdmin(admin.ModelAdmin):
    list_display    = ('certificate_number', 'customer', 'product_description',
                       'date_of_inspection', 'due_next_inspection', 'technician_name')
    search_fields   = ('customer__name', 'product_description', 'model_number', 'serial_number')
    list_filter     = ('date_of_inspection',)
    inlines         = [TestEquipmentRowInline, LoadCellCertRowInline]
    readonly_fields = ('certificate_number', 'created_at')


@admin.register(LoadCellSavedCert)
class LoadCellSavedCertAdmin(admin.ModelAdmin):
    list_display  = ('cert_number',)
    search_fields = ('cert_number',)


@admin.register(PhotoReference)
class PhotoReferenceAdmin(admin.ModelAdmin):
    list_display  = ('equipment_name', 'source_type', 'serial_number',
                     'part_number', 'record_date', 'uploaded_by', 'uploaded_at')
    search_fields = ('equipment_name', 'serial_number', 'part_number', 'caption')
    list_filter   = ('source_type', 'record_date')