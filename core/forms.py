from django import forms
from django.contrib.auth.models import User
from django.forms import inlineformset_factory
from .models import (Certificate, Customer, TestEquipmentRow, LoadCellCertRow,
                     LoadTest, Inventory, RepairInspection, PhotoReference)

W  = {'class': 'form-control'}
WD = {'class': 'form-control', 'type': 'date'}
WT = {'class': 'form-control', 'rows': 3}

def ac(field, model, extra=None):
    attrs = {
        'class': 'form-control',
        'data-autocomplete': 'true',
        'data-ac-field': field,
        'data-ac-model': model,
        'autocomplete': 'off',
    }
    if extra:
        attrs.update(extra)
    return attrs


class CustomerForm(forms.ModelForm):
    class Meta:
        model   = Customer
        fields  = ['name', 'address']
        widgets = {
            'name':    forms.TextInput(attrs=W),
            'address': forms.Textarea(attrs=WT),
        }


# ── LOAD TEST ────────────────────────────────────────────────────────────────

class LoadTestForm(forms.ModelForm):
    class Meta:
        model  = LoadTest
        fields = [
            'customer_name', 'customer_address',
            'description', 'capacity', 'model_number', 'serial_number',
            'equipment', 'tested_load', 'safe_load', 'certificate_number',
            'mechanic', 'date_inspection', 'date_due', 'remark',
        ]
        widgets = {
            'customer_name':      forms.TextInput(attrs=ac('customer_name', 'all', {'placeholder': 'e.g. Philippine Airlines'})),
            'customer_address':   forms.TextInput(attrs=ac('customer_address', 'all', {'placeholder': 'e.g. NAIA Terminal 2, Pasay City'})),
            'description':        forms.TextInput(attrs=ac('description', 'loadtest')),
            'capacity':           forms.TextInput(attrs=ac('capacity', 'loadtest', {'placeholder': 'e.g. 5,000 lbs'})),
            'model_number':       forms.TextInput(attrs=ac('model_number', 'all')),
            'serial_number':      forms.TextInput(attrs=ac('serial_number', 'all')),
            'equipment':          forms.TextInput(attrs=ac('equipment', 'loadtest')),
            'tested_load':        forms.TextInput(attrs=ac('tested_load', 'loadtest')),
            'safe_load':          forms.TextInput(attrs=ac('safe_load', 'loadtest')),
            'certificate_number': forms.TextInput(attrs=ac('certificate_number', 'loadtest')),
            'mechanic':           forms.TextInput(attrs=ac('mechanic', 'all')),
            'date_inspection':    forms.DateInput(attrs=WD),
            'date_due':           forms.DateInput(attrs=WD),
            'remark':             forms.Textarea(attrs=WT),
        }


# ── INVENTORY ────────────────────────────────────────────────────────────────

class InventoryForm(forms.ModelForm):
    class Meta:
        model  = Inventory
        fields = [
            'customer_name', 'customer_address',
            'brand', 'description', 'part_number',
            'location', 'quantity', 'less_by',
            'use', 'remarks',
        ]
        widgets = {
            'customer_name':    forms.TextInput(attrs=ac('customer_name', 'all', {'placeholder': 'e.g. Philippine Airlines'})),
            'customer_address': forms.TextInput(attrs=ac('customer_address', 'all', {'placeholder': 'e.g. NAIA Terminal 2, Pasay City'})),
            'brand':            forms.TextInput(attrs=ac('brand', 'inventory', {'placeholder': 'e.g. Snap-on'})),
            'description':      forms.TextInput(attrs=ac('description', 'inventory')),
            'part_number':      forms.TextInput(attrs=ac('part_number', 'inventory', {'placeholder': 'e.g. PN-98765'})),
            'location':         forms.TextInput(attrs=ac('location', 'inventory')),
            'quantity':         forms.NumberInput(attrs={**W, 'min': '0'}),
            'less_by':          forms.NumberInput(attrs={**W, 'min': '0', 'placeholder': '0'}),
            'use':              forms.Textarea(attrs={**W, 'rows': 2, 'placeholder': 'Describe the purpose or use of this item…'}),
            'remarks':          forms.Textarea(attrs=WT),
        }


# ── REPAIR / INSPECTION ───────────────────────────────────────────────────────

class RepairInspectionForm(forms.ModelForm):
    class Meta:
        model  = RepairInspection
        fields = [
            'customer_name', 'customer_address',
            'record_type', 'description', 'capacity',
            'model_number', 'serial_number',
            'report_number', 'mechanic', 'date',
            'customer_report', 'diagnose_result', 'remarks',
        ]
        widgets = {
            'customer_name':    forms.TextInput(attrs=ac('customer_name', 'all', {'placeholder': 'e.g. Philippine Airlines'})),
            'customer_address': forms.TextInput(attrs=ac('customer_address', 'all', {'placeholder': 'e.g. NAIA Terminal 2, Pasay City'})),
            'record_type':      forms.Select(attrs=W),
            'description':      forms.TextInput(attrs=ac('description', 'repair')),
            'capacity':         forms.TextInput(attrs=ac('capacity', 'repair', {'placeholder': 'e.g. 3,000 lbs'})),
            'model_number':     forms.TextInput(attrs=ac('model_number', 'all')),
            'serial_number':    forms.TextInput(attrs=ac('serial_number', 'all')),
            'report_number':    forms.TextInput(attrs=ac('report_number', 'repair')),
            'mechanic':         forms.TextInput(attrs=ac('mechanic', 'all')),
            'date':             forms.DateInput(attrs=WD),
            'customer_report':  forms.Textarea(attrs=WT),
            'diagnose_result':  forms.Textarea(attrs=WT),
            'remarks':          forms.Textarea(attrs=WT),
        }


# ── CERTIFICATE ──────────────────────────────────────────────────────────────

class CertificateForm(forms.ModelForm):
    class Meta:
        model  = Certificate
        fields = [
            'customer', 'address', 'accessories',
            'date_of_inspection', 'due_next_inspection',
            'product_description', 'capacity', 'model_number', 'serial_number',
            'working_load_limits', 'tested_load', 'pressure_relief',
            'load_cell_certificate_number',
            'work_perform',
            'technician_name', 'technician_date',
        ]
        widgets = {
            'customer':                      forms.Select(attrs={**W, 'id': 'id_customer'}),
            'address':                       forms.Textarea(attrs={**W, 'rows': 2, 'id': 'id_address'}),
            'accessories':                   forms.Textarea(attrs={**W, 'rows': 2}),
            'date_of_inspection':            forms.DateInput(attrs=WD),
            'due_next_inspection':           forms.DateInput(attrs=WD),
            'product_description':           forms.TextInput(attrs=W),
            'capacity':                      forms.TextInput(attrs=W),
            'model_number':                  forms.TextInput(attrs=W),
            'serial_number':                 forms.TextInput(attrs=W),
            'working_load_limits':           forms.TextInput(attrs=W),
            'tested_load':                   forms.TextInput(attrs=W),
            'pressure_relief':               forms.TextInput(attrs=W),
            'load_cell_certificate_number':  forms.TextInput(attrs=W),
            'work_perform':                  forms.Textarea(attrs={**W, 'rows': 5}),
            'technician_name':               forms.TextInput(attrs=ac('mechanic', 'all')),
            'technician_date':               forms.DateInput(attrs=WD),
        }


class TestEquipmentRowForm(forms.ModelForm):
    class Meta:
        model   = TestEquipmentRow
        fields  = ['description', 'capacity', 'model_no', 'serial_no', 'part_no', 'expiry_date']
        widgets = {f: forms.TextInput(attrs={'class': 'form-control tbl-input'}) for f in
                   ['description', 'capacity', 'model_no', 'serial_no', 'part_no', 'expiry_date']}


TestEquipmentFormSet = inlineformset_factory(
    Certificate, TestEquipmentRow,
    form=TestEquipmentRowForm,
    extra=3, can_delete=True,
    fields=['description', 'capacity', 'model_no', 'serial_no', 'part_no', 'expiry_date']
)


class LoadCellCertRowForm(forms.ModelForm):
    class Meta:
        model   = LoadCellCertRow
        fields  = ['cert_number']
        widgets = {
            'cert_number': forms.TextInput(attrs={
                'class': 'form-control tbl-input',
                'placeholder': 'e.g. LF-2026-0006',
                'autocomplete': 'off',
            })
        }


LoadCellCertFormSet = inlineformset_factory(
    Certificate, LoadCellCertRow,
    form=LoadCellCertRowForm,
    extra=1, can_delete=True,
    fields=['cert_number']
)


# ── PHOTO REFERENCE ──────────────────────────────────────────────────────────

class PhotoReferenceForm(forms.ModelForm):
    class Meta:
        model  = PhotoReference
        fields = [
            'photo', 'caption', 'source_type',
            'record_date', 'equipment_name', 'serial_number', 'part_number',
        ]
        widgets = {
            'photo':          forms.ClearableFileInput(attrs={'class': 'form-control', 'accept': 'image/*', 'capture': 'environment'}),
            'caption':        forms.TextInput(attrs={**W, 'placeholder': 'Optional caption…'}),
            'source_type':    forms.Select(attrs=W),
            'record_date':    forms.DateInput(attrs=WD),
            'equipment_name': forms.TextInput(attrs={**W, 'placeholder': 'e.g. Hydraulic Jack'}),
            'serial_number':  forms.TextInput(attrs={**W, 'placeholder': 'e.g. SN-12345'}),
            'part_number':    forms.TextInput(attrs={**W, 'placeholder': 'e.g. PN-98765'}),
        }


# ── USER / ACCOUNT MANAGEMENT ─────────────────────────────────────────────────

class CreateUserForm(forms.ModelForm):
    password = forms.CharField(
        required=False,
        widget=forms.PasswordInput(attrs={**W, 'placeholder': 'Leave blank to keep current password'}),
    )
    confirm_password = forms.CharField(
        required=False,
        widget=forms.PasswordInput(attrs={**W, 'placeholder': 'Repeat password'}),
        label='Confirm Password',
    )

    class Meta:
        model  = User
        fields = ['username', 'first_name', 'last_name', 'email', 'is_staff', 'is_active']
        widgets = {
            'username':   forms.TextInput(attrs=W),
            'first_name': forms.TextInput(attrs=W),
            'last_name':  forms.TextInput(attrs=W),
            'email':      forms.EmailInput(attrs=W),
        }
        help_texts = {'username': '', 'is_staff': 'Staff/Admin users can manage accounts.'}

    def clean(self):
        cleaned = super().clean()
        pw  = cleaned.get('password')
        cpw = cleaned.get('confirm_password')
        if not self.instance.pk and not pw:
            self.add_error('password', 'Password is required for new accounts.')
        if pw and pw != cpw:
            self.add_error('confirm_password', 'Passwords do not match.')
        return cleaned