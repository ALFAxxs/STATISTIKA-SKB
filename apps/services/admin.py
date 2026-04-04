# apps/services/admin.py

from django.contrib import admin
from .models import ServiceCategory, Service, PatientService


class ServiceInline(admin.TabularInline):
    model = Service
    extra = 1
    fields = ['code', 'name', 'price_normal', 'price_railway', 'department', 'is_active']


@admin.register(ServiceCategory)
class ServiceCategoryAdmin(admin.ModelAdmin):
    inlines = [ServiceInline]
    list_display = ['icon', 'name', 'code', 'category_type', 'is_active']
    list_filter = ['category_type', 'is_active']
    search_fields = ['name', 'code']
    list_editable = ['is_active']


@admin.register(Service)
class ServiceAdmin(admin.ModelAdmin):
    list_display = [
        'code', 'name', 'category',
        'price_normal', 'price_railway', 'department', 'is_active'
    ]
    list_filter = ['category', 'department', 'is_active']
    search_fields = ['name', 'code']
    list_editable = ['price_normal', 'price_railway', 'is_active']


@admin.register(PatientService)
class PatientServiceAdmin(admin.ModelAdmin):
    list_display = [
        'patient_card', 'service', 'quantity',
        'price', 'total_price_display',
        'status', 'is_paid', 'ordered_at'
    ]
    list_filter = ['status', 'is_paid', 'service__category', 'patient_category_at_order']
    search_fields = ['patient_card__full_name', 'patient_card__medical_record_number']
    readonly_fields = ['ordered_at', 'patient_category_at_order']

    def total_price_display(self, obj):
        return f"{obj.total_price:,.0f} so'm"
    total_price_display.short_description = "Jami narx"
from .models import Medicine, PatientMedicine

@admin.register(Medicine)
class MedicineAdmin(admin.ModelAdmin):
    list_display = ['name', 'unit', 'is_active']
    list_filter = ['unit', 'is_active']
    search_fields = ['name']

@admin.register(PatientMedicine)
class PatientMedicineAdmin(admin.ModelAdmin):
    list_display = ['patient_card', 'medicine', 'quantity', 'price', 'ordered_by', 'ordered_at']
    list_filter = ['medicine__unit']
    search_fields = ['medicine__name', 'patient_card__full_name']
