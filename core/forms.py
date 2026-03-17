from django import forms
from django.contrib.auth.models import User
from django.forms import inlineformset_factory
from .models import Certificate, Customer, TestEquipmentRow, LoadCellCertRow, LoadTest, Inventory, RepairInspection

W  = {'class': 'form-control'}
WD = {'class': 'form-control', 'type': 'date'}
WT = {'class': 'form-control', 'rows': 3}


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
            'company', 'description', 'model_number', 'serial_number',
            'equipment', 'tested_load', 'safe_load', 'certificate_number',
            'mechanic', 'date_inspection', 'date_due', 'remark',
        ]
        widgets = {
            'customer_name':      forms.TextInput(attrs={**W, 'placeholder': 'e.g. Philippine Airlines'}),
            'customer_address':   forms.TextInput(attrs={**W, 'placeholder': 'e.g. NAIA Terminal 2, Pasay City'}),
            'company':            forms.TextInput(attrs=W),
            'description':        forms.TextInput(attrs=W),
            'model_number':       forms.TextInput(attrs=W),
            'serial_number':      forms.TextInput(attrs=W),
            'equipment':          forms.TextInput(attrs=W),
            'tested_load':        forms.TextInput(attrs=W),
            'safe_load':          forms.TextInput(attrs=W),
            'certificate_number': forms.TextInput(attrs=W),
            'mechanic':           forms.TextInput(attrs=W),
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
            'description', 'model_number', 'serial_number',
            'location', 'quantity', 'remarks',
        ]
        widgets = {
            'customer_name':    forms.TextInput(attrs={**W, 'placeholder': 'e.g. Philippine Airlines'}),
            'customer_address': forms.TextInput(attrs={**W, 'placeholder': 'e.g. NAIA Terminal 2, Pasay City'}),
            'description':      forms.TextInput(attrs=W),
            'model_number':     forms.TextInput(attrs=W),
            'serial_number':    forms.TextInput(attrs=W),
            'location':         forms.TextInput(attrs=W),
            'quantity':         forms.NumberInput(attrs=W),
            'remarks':          forms.Textarea(attrs=WT),
        }


# ── REPAIR / INSPECTION ──────────────────────────────────────────────────────

class RepairInspectionForm(forms.ModelForm):
    class Meta:
        model  = RepairInspection
        fields = [
            'customer_name', 'customer_address',
            'record_type', 'company', 'description',
            'model_number', 'serial_number',
            'report_number', 'mechanic', 'date',
            'customer_report', 'diagnose_result', 'remarks',
        ]
        widgets = {
            'customer_name':    forms.TextInput(attrs={**W, 'placeholder': 'e.g. Philippine Airlines'}),
            'customer_address': forms.TextInput(attrs={**W, 'placeholder': 'e.g. NAIA Terminal 2, Pasay City'}),
            'record_type':      forms.Select(attrs=W),
            'company':          forms.TextInput(attrs=W),
            'description':      forms.TextInput(attrs=W),
            'model_number':     forms.TextInput(attrs=W),
            'serial_number':    forms.TextInput(attrs=W),
            'report_number':    forms.TextInput(attrs=W),
            'mechanic':         forms.TextInput(attrs=W),
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
            'technician_name':               forms.TextInput(attrs=W),
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


# ── LOAD CELL CERTIFICATE ROW ────────────────────────────────────────────────

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


# ── USER / ACCOUNT MANAGEMENT (ADMIN ONLY) ───────────────────────────────────

class CreateUserForm(forms.ModelForm):
    """
    Form for admin to create or edit a user account.
    Password is optional on edit — leave blank to keep existing password.
    """
    password = forms.CharField(
        required=False,
        widget=forms.PasswordInput(attrs={**W, 'placeholder': 'Leave blank to keep current password'}),
        help_text='Required when creating. Leave blank when editing to keep existing password.',
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
        help_texts = {
            'username': '',
            'is_staff': 'Staff/Admin users can manage accounts.',
        }

    def clean(self):
        cleaned = super().clean()
        pw  = cleaned.get('password')
        cpw = cleaned.get('confirm_password')

        # If this is a new user (no pk yet), password is required
        if not self.instance.pk and not pw:
            self.add_error('password', 'Password is required for new accounts.')

        if pw and pw != cpw:
            self.add_error('confirm_password', 'Passwords do not match.')

        return cleaned