from django.db import models
from django.contrib.auth.models import User


class Customer(models.Model):
    name    = models.CharField(max_length=255)
    address = models.TextField(blank=True)

    def __str__(self):
        return self.name


# ── LOAD TEST ────────────────────────────────────────────────────────────────

class LoadTest(models.Model):
    customer_name      = models.CharField(max_length=255)
    customer_address   = models.CharField(max_length=255, blank=True)
    company            = models.CharField(max_length=255, blank=True)
    description        = models.CharField(max_length=255, blank=True)
    capacity           = models.CharField(max_length=100, blank=True)
    model_number       = models.CharField(max_length=100, blank=True)
    serial_number      = models.CharField(max_length=100, blank=True)
    equipment          = models.CharField(max_length=255, blank=True)
    tested_load        = models.CharField(max_length=100, blank=True)
    safe_load          = models.CharField(max_length=100, blank=True)
    certificate_number = models.CharField(max_length=100, blank=True)
    mechanic           = models.CharField(max_length=255, blank=True)
    date_inspection    = models.DateField(null=True, blank=True)
    date_due           = models.DateField(null=True, blank=True)
    remark             = models.TextField(blank=True)
    created_at         = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.customer_name} — {self.description}"


# ── INVENTORY ────────────────────────────────────────────────────────────────

class Inventory(models.Model):
    customer_name    = models.CharField(max_length=255)
    customer_address = models.CharField(max_length=255, blank=True)
    brand            = models.CharField(max_length=255, blank=True)
    description      = models.CharField(max_length=255, blank=True)
    part_number      = models.CharField(max_length=100, blank=True)
    location         = models.CharField(max_length=255, blank=True)
    quantity         = models.PositiveIntegerField(default=0)
    less_by          = models.PositiveIntegerField(default=0, help_text='Amount to deduct from current stock')
    use              = models.TextField(blank=True, help_text='Purpose or usage of this item')
    remarks          = models.TextField(blank=True)
    created_at       = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.description} ({self.part_number})"

    @property
    def stock_on_hand(self):
        return max(self.quantity - self.less_by, 0)

    class Meta:
        verbose_name_plural = 'Inventories'


# ── REPAIR / INSPECTION ───────────────────────────────────────────────────────

class RepairInspection(models.Model):
    RECORD_TYPE_CHOICES = [
        ('repair',     'Repair'),
        ('inspection', 'Inspection'),
    ]

    customer_name    = models.CharField(max_length=255)
    customer_address = models.CharField(max_length=255, blank=True)
    company          = models.CharField(max_length=255, blank=True)
    record_type      = models.CharField(max_length=20, choices=RECORD_TYPE_CHOICES, default='repair')
    description      = models.CharField(max_length=255, blank=True)
    capacity         = models.CharField(max_length=100, blank=True)
    model_number     = models.CharField(max_length=100, blank=True)
    serial_number    = models.CharField(max_length=100, blank=True)
    report_number    = models.CharField(max_length=100, blank=True)
    mechanic         = models.CharField(max_length=255, blank=True)
    date             = models.DateField(null=True, blank=True)
    customer_report  = models.TextField(blank=True)
    diagnose_result  = models.TextField(blank=True)
    remarks          = models.TextField(blank=True)
    created_at       = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.get_record_type_display()} — {self.customer_name} ({self.report_number})"


# ── CERTIFICATE ──────────────────────────────────────────────────────────────

class Certificate(models.Model):
    customer                     = models.ForeignKey(Customer, on_delete=models.PROTECT)
    address                      = models.TextField(blank=True)
    accessories                  = models.TextField(blank=True)
    date_of_inspection           = models.DateField(null=True, blank=True)
    due_next_inspection          = models.DateField(null=True, blank=True)
    product_description          = models.CharField(max_length=255, blank=True)
    capacity                     = models.CharField(max_length=100, blank=True)
    model_number                 = models.CharField(max_length=100, blank=True)
    serial_number                = models.CharField(max_length=100, blank=True)
    working_load_limits          = models.CharField(max_length=100, blank=True)
    tested_load                  = models.CharField(max_length=100, blank=True)
    pressure_relief              = models.CharField(max_length=100, blank=True)
    load_cell_certificate_number = models.CharField(max_length=100, blank=True)
    work_perform                 = models.TextField(blank=True)
    technician_name              = models.CharField(max_length=255, blank=True)
    technician_date              = models.DateField(null=True, blank=True)
    created_at                   = models.DateTimeField(auto_now_add=True)

    @property
    def certificate_number(self):
        return f"CERT-{self.pk:04d}"

    def __str__(self):
        return f"{self.certificate_number} — {self.customer.name}"


class TestEquipmentRow(models.Model):
    certificate = models.ForeignKey(Certificate, on_delete=models.CASCADE, related_name='test_equipment_rows')
    description = models.CharField(max_length=255, blank=True)
    capacity    = models.CharField(max_length=100, blank=True)
    model_no    = models.CharField(max_length=100, blank=True)
    serial_no   = models.CharField(max_length=100, blank=True)
    part_no     = models.CharField(max_length=100, blank=True)
    expiry_date = models.CharField(max_length=100, blank=True)

    def __str__(self):
        return f"{self.description} ({self.serial_no})"


class LoadCellCertRow(models.Model):
    certificate = models.ForeignKey(Certificate, on_delete=models.CASCADE, related_name='load_cell_cert_rows')
    cert_number = models.CharField(max_length=100, blank=True)

    def __str__(self):
        return self.cert_number


# ── LOAD CELL SAVED CERT (reusable cert number library) ──────────────────────

class LoadCellSavedCert(models.Model):
    cert_number = models.CharField(max_length=100, unique=True)

    def __str__(self):
        return self.cert_number

    class Meta:
        ordering = ['cert_number']


# ── PHOTO REFERENCE ──────────────────────────────────────────────────────────

class PhotoReference(models.Model):
    SOURCE_CHOICES = [
        ('loadtest',   'Load Test'),
        ('repair',     'Repair'),
        ('inspection', 'Inspection'),
        ('inventory',  'Inventory'),
        ('general',    'General'),
    ]

    photo          = models.ImageField(upload_to='photos/%Y/%m/')
    caption        = models.CharField(max_length=255, blank=True)
    source_type    = models.CharField(max_length=20, choices=SOURCE_CHOICES, default='general')
    record_date    = models.DateField(null=True, blank=True)
    equipment_name = models.CharField(max_length=255, blank=True)
    serial_number  = models.CharField(max_length=100, blank=True)
    part_number    = models.CharField(max_length=100, blank=True)
    uploaded_by    = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    uploaded_at    = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.equipment_name or 'Photo'} — {self.uploaded_at:%Y-%m-%d}"