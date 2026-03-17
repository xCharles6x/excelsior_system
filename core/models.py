from django.db import models
from django.conf import settings
import datetime


class Customer(models.Model):
    name    = models.CharField(max_length=255)
    address = models.TextField(blank=True)
    created = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


class LoadTest(models.Model):
    customer_name    = models.CharField(max_length=255, blank=True)
    customer_address = models.CharField(max_length=500, blank=True)
    company          = models.CharField(max_length=255, blank=True)
    description      = models.CharField(max_length=255, blank=True)
    model_number     = models.CharField(max_length=100, blank=True)
    serial_number    = models.CharField(max_length=100, blank=True)
    equipment        = models.CharField(max_length=255, blank=True)
    tested_load      = models.CharField(max_length=100, blank=True)
    safe_load        = models.CharField(max_length=100, blank=True)
    certificate_number = models.CharField(max_length=100, blank=True)
    mechanic         = models.CharField(max_length=150, blank=True)
    date_inspection  = models.DateField(null=True, blank=True)
    date_due         = models.DateField(null=True, blank=True)
    remark           = models.TextField(blank=True)
    created_at       = models.DateTimeField(auto_now_add=True)
    created_by       = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL, related_name='loadtest_records')

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f"{self.company} – {self.description} ({self.date_inspection})"


class Inventory(models.Model):
    customer_name    = models.CharField(max_length=255, blank=True)
    customer_address = models.CharField(max_length=500, blank=True)
    description      = models.CharField(max_length=255, blank=True)
    model_number     = models.CharField(max_length=100, blank=True)
    serial_number    = models.CharField(max_length=100, blank=True)
    location         = models.CharField(max_length=255, blank=True)
    quantity         = models.IntegerField(default=0)
    remarks          = models.TextField(blank=True)
    created_at       = models.DateTimeField(auto_now_add=True)
    created_by       = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL, related_name='inventory_records')

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f"{self.description} ({self.serial_number})"


class RepairInspection(models.Model):
    RECORD_TYPE_CHOICES = [
        ('repair',     'Repair'),
        ('inspection', 'Inspection'),
    ]
    customer_name    = models.CharField(max_length=255, blank=True)
    customer_address = models.CharField(max_length=500, blank=True)
    record_type      = models.CharField(max_length=20, choices=RECORD_TYPE_CHOICES, default='repair')
    company          = models.CharField(max_length=255, blank=True)
    description      = models.CharField(max_length=255, blank=True)
    model_number     = models.CharField(max_length=100, blank=True)
    serial_number    = models.CharField(max_length=100, blank=True)
    report_number    = models.CharField(max_length=100, blank=True)
    mechanic         = models.CharField(max_length=150, blank=True)
    date             = models.DateField(null=True, blank=True)
    customer_report  = models.TextField(blank=True)
    diagnose_result  = models.TextField(blank=True)
    remarks          = models.TextField(blank=True)
    created_at       = models.DateTimeField(auto_now_add=True)
    created_by       = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL, related_name='repair_records')

    class Meta:
        ordering = ['-created_at']

    def get_record_type_display(self):
        return dict(self.RECORD_TYPE_CHOICES).get(self.record_type, self.record_type)

    def __str__(self):
        return f"{self.get_record_type_display()} – {self.company} ({self.date})"


class Certificate(models.Model):
    certificate_number  = models.CharField(max_length=50, unique=True, editable=False)
    created_at          = models.DateTimeField(auto_now_add=True)
    updated_at          = models.DateTimeField(auto_now=True)
    customer            = models.ForeignKey(Customer, on_delete=models.PROTECT, related_name='certificates')
    address             = models.TextField(blank=True)
    accessories         = models.TextField(blank=True)
    date_of_inspection  = models.DateField()
    due_next_inspection = models.DateField()
    product_description = models.CharField(max_length=255)
    capacity            = models.CharField(max_length=100, blank=True)
    model_number        = models.CharField(max_length=100, blank=True)
    serial_number       = models.CharField(max_length=100, blank=True)
    working_load_limits = models.CharField(max_length=100, blank=True)
    tested_load         = models.CharField(max_length=100, blank=True)
    pressure_relief     = models.CharField(max_length=100, blank=True)
    load_cell_certificate_number = models.CharField(max_length=100, blank=True)
    work_perform = models.TextField(blank=True, default=(
        "The described equipment was set up and tested in accordance with Tronair "
        "and Service Manual. The equipment should only be used in configuration as "
        "tested with the total load less than or equal to the rated working load limit.\n\n"
        "I hereby certify that the above-described equipment is compatible to meeting "
        "the load ratings when used under normal and proper use."
    ))
    technician_name = models.CharField(max_length=150)
    technician_date = models.DateField()

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return self.certificate_number

    def save(self, *args, **kwargs):
        if not self.certificate_number:
            self.certificate_number = self._generate_cert_number()
        if not self.address and self.customer_id:
            self.address = self.customer.address
        super().save(*args, **kwargs)

    @staticmethod
    def _generate_cert_number():
        year = datetime.date.today().year
        # Get the highest sequence number across ALL certificates regardless of year
        highest = 10286  # one below the desired start
        for cert in Certificate.objects.all():
            try:
                seq = int(cert.certificate_number.split('-')[1])
                if seq > highest:
                    highest = seq
            except (IndexError, ValueError):
                pass
        return f'{year}-{highest + 1}'


class TestEquipmentRow(models.Model):
    certificate = models.ForeignKey(Certificate, on_delete=models.CASCADE, related_name='test_equipment_rows')
    description = models.CharField(max_length=255, blank=True)
    capacity    = models.CharField(max_length=100, blank=True)
    model_no    = models.CharField(max_length=100, blank=True)
    serial_no   = models.CharField(max_length=100, blank=True)
    part_no     = models.CharField(max_length=100, blank=True)
    expiry_date = models.CharField(max_length=100, blank=True)
    order       = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order']

    def __str__(self):
        return f"{self.description} ({self.serial_no})"


# ── LOAD CELL SAVED LIST ─────────────────────────────────────────────────────

class LoadCellSavedCert(models.Model):
    cert_number = models.CharField(max_length=255, unique=True)
    created_at  = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['cert_number']

    def __str__(self):
        return self.cert_number


# ── LOAD CELL CERTIFICATE ROW ─────────────────────────────────────────────

class LoadCellCertRow(models.Model):
    certificate = models.ForeignKey(Certificate, on_delete=models.CASCADE, related_name='load_cell_rows')
    cert_number = models.CharField(max_length=255, blank=True, verbose_name='Load Cell Certificate Number')
    order       = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order']

    def __str__(self):
        return self.cert_number