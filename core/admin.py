from django.contrib import admin
from .models import Customer, LoadTest, Inventory, RepairInspection, Certificate, TestEquipmentRow, LoadCellCertRow


@admin.register(Customer)
class CustomerAdmin(admin.ModelAdmin):
    list_display  = ('name', 'address', 'created')
    search_fields = ('name', 'address')


@admin.register(LoadTest)
class LoadTestAdmin(admin.ModelAdmin):
    list_display  = ('customer_name', 'company', 'description', 'serial_number',
                     'certificate_number', 'mechanic', 'date_inspection', 'date_due', 'created_by', 'created_at')
    search_fields = ('customer_name', 'company', 'description', 'serial_number',
                     'certificate_number', 'mechanic')
    list_filter   = ('date_inspection',)


@admin.register(Inventory)
class InventoryAdmin(admin.ModelAdmin):
    list_display  = ('customer_name', 'description', 'model_number', 'serial_number',
                     'location', 'quantity', 'created_by', 'created_at')
    search_fields = ('customer_name', 'description', 'model_number', 'serial_number', 'location')


@admin.register(RepairInspection)
class RepairInspectionAdmin(admin.ModelAdmin):
    list_display  = ('customer_name', 'record_type', 'company', 'description',
                     'serial_number', 'report_number', 'mechanic', 'date', 'created_by', 'created_at')
    search_fields = ('customer_name', 'company', 'description', 'serial_number',
                     'report_number', 'mechanic')
    list_filter   = ('record_type', 'date')


class TestEquipmentRowInline(admin.TabularInline):
    model  = TestEquipmentRow
    extra  = 1
    fields = ('description', 'capacity', 'model_no', 'serial_no', 'part_no', 'expiry_date', 'order')


class LoadCellCertRowInline(admin.TabularInline):
    model  = LoadCellCertRow
    extra  = 1
    fields = ('cert_number', 'order')


@admin.register(Certificate)
class CertificateAdmin(admin.ModelAdmin):
    list_display  = ('certificate_number', 'customer', 'product_description',
                     'date_of_inspection', 'due_next_inspection', 'technician_name')
    search_fields = ('certificate_number', 'customer__name', 'product_description',
                     'model_number', 'serial_number')
    list_filter   = ('date_of_inspection',)
    inlines       = [TestEquipmentRowInline, LoadCellCertRowInline]
    readonly_fields = ('certificate_number', 'created_at', 'updated_at')