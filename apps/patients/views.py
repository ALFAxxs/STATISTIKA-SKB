# apps/patients/views.py

from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.core.paginator import Paginator
from django.db.models import Q
from django.http import JsonResponse, HttpResponse
from django.contrib.auth.decorators import login_required
from apps.users.decorators import role_required, department_filter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io
import json
import uuid

from django.utils import timezone

from .forms import PatientCardForm, DeathCauseForm, SurgicalOperationFormSet, ReceptionForm
from .models import (
    PatientCard, ICD10Code, DischargeConclusion,
    Region, District, City, Village, Country, OperationType
)


# ==================== AJAX VIEWS ====================

@login_required
def add_conclusion(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            name = data.get('name', '').strip()
            if not name:
                return JsonResponse({'success': False, 'error': "Nom bo'sh"})
            obj, created = DischargeConclusion.objects.get_or_create(name=name)
            return JsonResponse({
                'success': True,
                'id': obj.id,
                'name': obj.name,
                'created': created
            })
        except Exception as e:
            return JsonResponse({'success': False, 'error': str(e)})
    return JsonResponse({'success': False, 'error': 'Faqat POST'})


def get_regions(request):
    country_id = request.GET.get('country_id')
    regions = Region.objects.filter(country_id=country_id).values('id', 'name')
    return JsonResponse(list(regions), safe=False)


def get_districts(request):
    region_id = request.GET.get('region_id')
    districts = District.objects.filter(region_id=region_id).values('id', 'name')
    return JsonResponse(list(districts), safe=False)


def get_cities(request):
    district_id = request.GET.get('district_id')
    cities = City.objects.filter(district_id=district_id).values('id', 'name')
    return JsonResponse(list(cities), safe=False)


def get_villages(request):
    district_id = request.GET.get('district_id')
    villages = Village.objects.filter(district_id=district_id).values('id', 'name')
    return JsonResponse(list(villages), safe=False)


def icd10_search(request):
    q = request.GET.get('q', '')
    if len(q) < 2:
        return JsonResponse([], safe=False)
    results = ICD10Code.objects.filter(
        Q(code__icontains=q) | Q(title_uz__icontains=q)
    )[:15]
    data = [{'code': r.code, 'title': r.title_uz} for r in results]
    return JsonResponse(data, safe=False)


def operation_type_search(request):
    q = request.GET.get('q', '')
    if len(q) < 1:
        results = OperationType.objects.filter(is_active=True)[:20]
    else:
        results = OperationType.objects.filter(
            Q(name__icontains=q) | Q(code__icontains=q),
            is_active=True
        )[:15]
    data = [{'id': r.id, 'name': str(r)} for r in results]
    return JsonResponse(data, safe=False)


# ==================== PDF ====================

@login_required
def patient_card_pdf(request, pk):
    patient = get_object_or_404(
        PatientCard.objects.select_related(
            'department', 'attending_doctor', 'department_head',
            'referral_organization', 'country', 'region', 'district',
            'city', 'discharge_conclusion'
        ).prefetch_related('operations__operation_type'),
        pk=pk
    )

    # Bo'lim tekshiruvi
    if not request.user.is_superuser and request.user.role != 'admin':
        if request.user.role == 'reception':
            if patient.registered_by != request.user:
                messages.error(request, "Ruxsat yo'q.")
                return redirect('patient_list')
        elif request.user.department and patient.department != request.user.department:
            messages.error(request, "Ruxsat yo'q.")
            return redirect('patient_list')

    death_cause = getattr(patient, 'death_cause', None)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=2*cm, leftMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'title', parent=styles['Heading1'],
        fontSize=13, alignment=1, spaceAfter=8
    )
    bold_style = ParagraphStyle(
        'bold', parent=styles['Normal'],
        fontSize=9, fontName='Helvetica-Bold'
    )
    normal_style = ParagraphStyle('normal', parent=styles['Normal'], fontSize=9)

    elements = []
    elements.append(Paragraph(
        "SHIFOXONADAN CHIQARILGAN BEMOR STATISTIK KARTASI", title_style
    ))
    elements.append(Spacer(1, 0.3*cm))

    def row(label, value):
        return [
            Paragraph(str(label), bold_style),
            Paragraph(str(value) if value else '—', normal_style)
        ]

    address_parts = [
        str(patient.country) if patient.country else '',
        str(patient.region) if patient.region else '',
        str(patient.district) if patient.district else '',
        str(patient.city) if patient.city else '',
        patient.street_address or '',
    ]
    full_address = ', '.join(filter(None, address_parts)) or '—'

    data = [
        [Paragraph("BEMOR MA'LUMOTLARI", bold_style), ''],
        row("Tibbiy bayonnoma raqami:", patient.medical_record_number),
        row("Ism-familiya:", patient.full_name),
        row("Rezidentlik:", patient.get_resident_status_display()),
        row("Bemor kategoriyasi:", patient.get_patient_category_display()),
        row("Jinsi:", patient.get_gender_display()),
        row("Tug'ilgan sana:", patient.birth_date.strftime('%d.%m.%Y') if patient.birth_date else ''),
        row("Telefon:", patient.phone or '—'),
        row("Manzil:", full_address),
        row("Ijtimoiy holat:", patient.get_social_status_display() if patient.social_status else '—'),
        row("Ish joyi:", patient.workplace or '—'),
        row("Passport:", patient.passport_serial or '—'),
        row("JSHSHIR:", patient.JSHSHIR or '—'),

        [Paragraph("QABUL MA'LUMOTLARI", bold_style), ''],
        row("Kim olib kelgan:", patient.get_referral_type_display() if patient.referral_type else '—'),
        row("Yo'llagan muassasa:", str(patient.referral_organization) if patient.referral_organization else '—'),
        row("Yo'llagan muassasa tashxisi:", patient.referring_diagnosis or '—'),
        row("Qabul bo'limi tashxisi:", patient.admission_diagnosis or '—'),
        row("Kasallanishdan keyin:", patient.get_hours_after_illness_display() if patient.hours_after_illness else '—'),
        row("Shoshilinch:", 'Ha' if patient.is_emergency else "Yo'q"),
        row("Pullik:", 'Ha' if patient.is_paid else "Yo'q"),
        row("Shifoxona turi:", str(patient.hospital_type) if patient.hospital_type else '—'),
        row("Yotqizilgan sana:", patient.admission_date.strftime('%d.%m.%Y %H:%M') if patient.admission_date else ''),
        row("Bo'lim:", str(patient.department) if patient.department else '—'),
        row("Yotqizilish:", patient.get_admission_count_display() if patient.admission_count else '—'),

        [Paragraph("CHIQISH MA'LUMOTLARI", bold_style), ''],
        row("Yotgan kunlar:", f"{patient.days_in_hospital} kun"),
        row("Yakun:", patient.get_outcome_display() if patient.outcome else '—'),
        row("Chiqish xulosasi:", str(patient.discharge_conclusion) if patient.discharge_conclusion else '—'),
        row("Chiqgan sana:", patient.discharge_date.strftime('%d.%m.%Y %H:%M') if patient.discharge_date else '—'),

        [Paragraph("YAKUNIY TASHXIS", bold_style), ''],
        row("Klinik tashxis (MKB-10):", f"{patient.clinical_main_diagnosis or ''} {patient.clinical_main_diagnosis_text or ''}".strip() or '—'),
        row("Klinik yo'ldosh kasalliklar:", patient.clinical_comorbidities or '—'),
    ]

    if patient.outcome == 'deceased':
        data += [
            row("Patologoanatomik tashxis (MKB-10):", f"{patient.pathological_main_diagnosis or ''} {patient.pathological_main_diagnosis_text or ''}".strip() or '—'),
            row("Patologoanatomik yo'ldosh kasalliklar:", patient.pathological_comorbidities or '—'),
        ]
        if death_cause:
            data += [
                [Paragraph("O'LIM SABABI", bold_style), ''],
                row("a) Bevosita sabab:", death_cause.immediate_cause),
                row("b) Chaqiruvchi kasallik:", death_cause.underlying_cause),
                row("v) Asosiy kasallik kodi:", death_cause.main_disease_code),
                row("Boshqa muhim kasalliklar:", death_cause.other_significant_conditions or '—'),
            ]

    data += [
        [Paragraph("TEKSHIRUVLAR", bold_style), ''],
        row("OITS tekshiruvi:", f"{patient.aids_test_date or '—'} | {patient.aids_test_result or '—'}"),
        row("WP tekshiruvi:", f"{patient.wp_test_date or '—'} | {patient.wp_test_result or '—'}"),
        row("Urush qatnashchisi:", 'Ha' if patient.is_war_veteran else "Yo'q"),

        [Paragraph("SHIFOKORLAR", bold_style), ''],
        row("Davolovchi shifokor:", str(patient.attending_doctor) if patient.attending_doctor else '—'),
        row("Bo'lim mudiri:", str(patient.department_head) if patient.department_head else '—'),
    ]

    operations = patient.operations.all()
    if operations:
        data.append([Paragraph("JARROHLIK AMALIYOTLARI", bold_style), ''])
        for op in operations:
            name = str(op.operation_type) if op.operation_type else op.operation_name or '—'
            data.append(row(
                op.operation_date.strftime('%d.%m.%Y') if op.operation_date else '—',
                f"{name} | Narkoz: {op.get_anesthesia_display() if op.anesthesia else '—'} | Asorat: {op.complication or '—'}"
            ))

    table = Table(data, colWidths=[7*cm, 11*cm])
    table.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
        ('BACKGROUND', (0, 0), (1, 0), colors.HexColor('#1F4E79')),
        ('TEXTCOLOR', (0, 0), (1, 0), colors.white),
        ('SPAN', (0, 0), (1, 0)),
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)

    response = HttpResponse(buffer, content_type='application/pdf')
    response['Content-Disposition'] = f'inline; filename="bemor_{patient.medical_record_number}.pdf"'
    return response


# ==================== PATIENT VIEWS ====================

# apps/patients/views.py — patient_list view

@login_required
def patient_list(request):
    qs = PatientCard.objects.select_related(
        'department', 'attending_doctor'
    ).order_by('-admission_date')

    # Bo'lim filteri
    qs = department_filter(qs, request.user)

    # Qabulxona faqat o'zi registratsiya qilganlarni ko'radi
    if request.user.role == 'reception':
        qs = qs.filter(registered_by=request.user)

    # Qidiruv
    query = request.GET.get('q', '')
    if query:
        qs = qs.filter(
            Q(full_name__icontains=query) |
            Q(medical_record_number__icontains=query) |
            Q(passport_serial__icontains=query)
        )

    # Status filteri
    status = request.GET.get('status', '')
    if status:
        qs = qs.filter(status=status)

    # Yakun filteri
    outcome = request.GET.get('outcome', '')
    if outcome:
        qs = qs.filter(outcome=outcome)

    # Bo'lim filteri (faqat admin)
    dept_filter = request.GET.get('department', '')
    if dept_filter and (request.user.is_superuser or request.user.role == 'admin'):
        qs = qs.filter(department_id=dept_filter)

    # Shifokor filteri
    doctor_filter = request.GET.get('doctor', '')
    if doctor_filter:
        qs = qs.filter(attending_doctor_id=doctor_filter)

    # Filter uchun ro'yxatlar
    from .models import Department, Doctor
    if request.user.is_superuser or request.user.role == 'admin':
        departments = Department.objects.filter(is_active=True)
    else:
        departments = Department.objects.filter(
            pk=request.user.department.pk
        ) if request.user.department else Department.objects.none()

    doctors = Doctor.objects.filter(is_active=True).select_related('department')
    if not request.user.is_superuser and request.user.role != 'admin':
        if request.user.department:
            doctors = doctors.filter(department=request.user.department)

    paginator = Paginator(qs, 20)
    page = paginator.get_page(request.GET.get('page'))

    return render(request, 'patients/patient_list.html', {
        'page_obj': page,
        'query': query,
        'selected_status': status,
        'selected_outcome': outcome,
        'selected_dept': dept_filter,
        'selected_doctor': doctor_filter,
        'departments': departments,
        'doctors': doctors,
    })


@login_required
def patient_detail(request, pk):
    patient = get_object_or_404(
        PatientCard.objects.select_related(
            'department', 'attending_doctor', 'department_head',
            'referral_organization', 'country', 'region', 'district',
            'city', 'discharge_conclusion'
        ).prefetch_related('operations__operation_type'),
        pk=pk
    )

    # Bo'lim tekshiruvi
    if not request.user.is_superuser and request.user.role != 'admin':
        if request.user.role == 'reception':
            if patient.registered_by != request.user:
                messages.error(request, "Siz bu bemorni ko'rishga ruxsatingiz yo'q.")
                return redirect('patient_list')
        elif request.user.department and patient.department != request.user.department:
            messages.error(request, "Siz bu bemorni ko'rishga ruxsatingiz yo'q.")
            return redirect('patient_list')

    death_cause = getattr(patient, 'death_cause', None)
    address_parts = filter(None, [
        str(patient.country) if patient.country else '',
        str(patient.region) if patient.region else '',
        str(patient.district) if patient.district else '',
        str(patient.city) if patient.city else '',
        patient.street_address or '',
    ])
    full_address = ', '.join(address_parts) or '—'

    return render(request, 'patients/patient_detail.html', {
        'patient': patient,
        'death_cause': death_cause,
        'full_address': full_address,
    })


@login_required
@role_required('admin', 'doctor', 'statistician')
def patient_card_create(request):
    if request.method == 'POST':
        form = PatientCardForm(request.POST)
        death_form = DeathCauseForm(request.POST)
        surgery_formset = SurgicalOperationFormSet(request.POST)

        is_deceased = request.POST.get('outcome') == 'deceased'
        forms_valid = form.is_valid() and surgery_formset.is_valid()
        if is_deceased:
            forms_valid = forms_valid and death_form.is_valid()

        if forms_valid:
            patient = form.save(commit=False)
            if not request.user.is_superuser and request.user.role != 'admin':
                if request.user.department:
                    patient.department = request.user.department
            patient.save()
            form.save_m2m()

            surgeries = surgery_formset.save(commit=False)
            for s in surgeries:
                s.patient_card = patient
                s.save()
            for s in surgery_formset.deleted_objects:
                s.delete()

            if is_deceased:
                death = death_form.save(commit=False)
                death.patient_card = patient
                death.save()

            messages.success(request, "Bemor kartasi saqlandi!")
            return redirect('patient_list')
        else:
            messages.error(request, "Formada xatoliklar bor. Tekshiring.")
    else:
        form = PatientCardForm()
        if not request.user.is_superuser and request.user.role != 'admin':
            if request.user.department:
                form.initial['department'] = request.user.department
        death_form = DeathCauseForm()
        surgery_formset = SurgicalOperationFormSet()

    return render(request, 'patients/patient_form.html', {
        'form': form,
        'death_form': death_form,
        'surgery_formset': surgery_formset,
        'title': 'Yangi bemor kartasi',
    })


@login_required
@role_required('admin', 'doctor', 'statistician', 'reception')
def patient_card_edit(request, pk):
    patient = get_object_or_404(PatientCard, pk=pk)

    # Ruxsat tekshiruvi
    if not request.user.is_superuser and request.user.role != 'admin':
        if request.user.role == 'reception':
            if patient.registered_by != request.user:
                messages.error(request, "Siz bu bemorni tahrirlay olmaysiz.")
                return redirect('patient_list')
        elif request.user.department and patient.department != request.user.department:
            messages.error(request, "Siz bu bemorni tahrirlay olmaysiz.")
            return redirect('patient_list')

    # Qabulxona faqat ReceptionForm ishlatadi
    is_reception = request.user.role == 'reception'
    FormClass = ReceptionForm if is_reception else PatientCardForm

    death_instance = getattr(patient, 'death_cause', None)

    if request.method == 'POST':
        form = FormClass(request.POST, instance=patient)

        if is_reception:
            if form.is_valid():
                form.save()
                messages.success(request, "Ma'lumotlar yangilandi!")
                return redirect('patient_list')
            else:
                messages.error(request, "Formada xatoliklar bor.")
        else:
            death_form = DeathCauseForm(request.POST, instance=death_instance)
            surgery_formset = SurgicalOperationFormSet(request.POST, instance=patient)

            is_deceased = request.POST.get('outcome') == 'deceased'
            forms_valid = form.is_valid() and surgery_formset.is_valid()
            if is_deceased:
                forms_valid = forms_valid and death_form.is_valid()

            if forms_valid:
                patient = form.save()
                surgeries = surgery_formset.save(commit=False)
                for s in surgeries:
                    s.patient_card = patient
                    s.save()
                for s in surgery_formset.deleted_objects:
                    s.delete()

                if is_deceased:
                    death = death_form.save(commit=False)
                    death.patient_card = patient
                    death.save()
                elif death_instance:
                    death_instance.delete()

                messages.success(request, "Bemor kartasi yangilandi!")
                return redirect('patient_list')
            else:
                messages.error(request, "Formada xatoliklar bor.")
    else:
        form = FormClass(instance=patient)
        if not is_reception:
            death_form = DeathCauseForm(instance=death_instance)
            surgery_formset = SurgicalOperationFormSet(instance=patient)

    if is_reception:
        return render(request, 'patients/reception_form.html', {
            'form': form,
            'title': f"Tahrirlash: {patient.full_name}",
            'patient': patient,
        })

    return render(request, 'patients/patient_form.html', {
        'form': form,
        'death_form': death_form,
        'surgery_formset': surgery_formset,
        'title': f"Tahrirlash: {patient.full_name}",
        'patient': patient,
    })


@login_required
@role_required('admin')
def patient_delete(request, pk):
    patient = get_object_or_404(PatientCard, pk=pk)
    if request.method == 'POST':
        patient.delete()
        messages.success(request, "Bemor kartasi o'chirildi.")
        return redirect('patient_list')
    return render(request, 'patients/patient_confirm_delete.html', {'patient': patient})


@login_required
@role_required('admin', 'reception')
def reception_create(request):
    if request.method == 'POST':
        form = ReceptionForm(request.POST)
        if form.is_valid():
            patient = form.save(commit=False)

            # Bo'limni avtomatik qo'yish
            if not request.user.is_superuser and request.user.role != 'admin':
                if request.user.department:
                    patient.department = request.user.department

            # Avtomatik bayonnoma raqami
            year = timezone.now().year
            while True:
                record_number = f"{year}-{str(uuid.uuid4())[:6].upper()}"
                if not PatientCard.objects.filter(
                    medical_record_number=record_number
                ).exists():
                    break

            patient.medical_record_number = record_number
            patient.status = 'registered'
            patient.registered_by = request.user
            patient.save()

            messages.success(
                request,
                f"✅ Bemor qabul qilindi! Bayonnoma: {patient.medical_record_number}"
            )
            return redirect('patient_list')
        else:
            messages.error(request, "Formada xatoliklar bor.")
    else:
        form = ReceptionForm()
        if request.user.department:
            form.initial['department'] = request.user.department

    return render(request, 'patients/reception_form.html', {
        'form': form,
        'title': 'Bemor qabul qilish',
    })