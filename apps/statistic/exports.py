# apps/statistic/exports.py

import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from django.db.models import Avg
from django.http import HttpResponse
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from apps.patients.models import PatientCard


def get_filtered_queryset(request):
    from apps.users.decorators import department_filter

    qs = PatientCard.objects.select_related(
        'department', 'attending_doctor', 'referral_organization',
        'country', 'region', 'district', 'city',
        'discharge_conclusion'
    ).prefetch_related(
        'operations__operation_type'
    ).order_by('admission_date')

    # Rol bo'yicha cheklash
    qs = department_filter(qs, request.user)

    # Barcha filterlar
    year = request.GET.get('year')
    month = request.GET.get('month')
    dept = request.GET.get('department')
    doctor = request.GET.get('doctor')
    outcome = request.GET.get('outcome')
    status = request.GET.get('status')
    gender = request.GET.get('gender')
    patient_category = request.GET.get('patient_category')
    resident_status = request.GET.get('resident_status')
    referral_type = request.GET.get('referral_type')
    date_from = request.GET.get('date_from')
    date_to = request.GET.get('date_to')

    if year:
        qs = qs.filter(admission_date__year=year)
    if month:
        qs = qs.filter(admission_date__month=month)
    if dept:
        qs = qs.filter(department_id=dept)
    if doctor:
        qs = qs.filter(attending_doctor_id=doctor)
    if outcome:
        qs = qs.filter(outcome=outcome)
    if status:
        qs = qs.filter(status=status)
    if gender:
        qs = qs.filter(gender=gender)
    if patient_category:
        qs = qs.filter(patient_category=patient_category)
    if resident_status:
        qs = qs.filter(resident_status=resident_status)
    if referral_type:
        qs = qs.filter(referral_type=referral_type)
    if date_from:
        qs = qs.filter(admission_date__date__gte=date_from)
    if date_to:
        qs = qs.filter(admission_date__date__lte=date_to)
    org_id = request.GET.get('org')
    if org_id:
        qs = qs.filter(workplace_org_id=org_id)

    return qs


def get_stats_queryset(request):
    """Statistika uchun alohida queryset — prefetch yo'q"""
    qs = PatientCard.objects.all()
    year = request.GET.get('year')
    dept = request.GET.get('department')
    if year:
        qs = qs.filter(admission_date__year=year)
    if dept:
        qs = qs.filter(department_id=dept)
    return qs


def get_full_address(patient):
    parts = filter(None, [
        str(patient.country) if patient.country else '',
        str(patient.region) if patient.region else '',
        str(patient.district) if patient.district else '',
        str(patient.city) if patient.city else '',
        patient.street_address or '',
    ])
    return ', '.join(parts) or '—'


def export_excel(request):
    qs = get_filtered_queryset(request)
    qs_stats = get_stats_queryset(request)

    wb = openpyxl.Workbook()

    # ==================== STIL O'ZGARUVCHILARI ====================
    header_font = Font(bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # ==================== 1-SAHIFA: BEMORLAR RO'YXATI ====================
    ws = wb.active
    ws.title = "Bemorlar ro'yxati"

    headers = [
        "№", "Bayonnoma №", "Ism-familiya", "Jinsi", "Tug'ilgan sana",
        "Rezident", "Manzil", "Ijtimoiy holat", "Passport", "Qabul turi",
        "Yo'llagan muassasa", "Qabul bo'limi tashxisi", "Necha soat keyin",
        "Shoshilinch", "Pullik", "Yotqizilgan sana", "Bo'lim",
        "Qayta/Birinchi", "Yotgan kunlar", "Yakun", "Chiqish xulosasi",
        "Chiqgan sana", "Klinik tashxis (MKB-10)", "Shifokor",
        "Jarrohlik amaliyotlari",
    ]

    ws.row_dimensions[1].height = 40
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    for row_num, patient in enumerate(qs, 2):
        # Jarrohlik amaliyotlarini yig'ish
        operations = patient.operations.all()
        if operations:
            op_lines = []
            for i, op in enumerate(operations, 1):
                name = str(op.operation_type) if op.operation_type else op.operation_name or '—'
                date = op.operation_date.strftime('%d.%m.%Y') if op.operation_date else '—'
                anesthesia = op.get_anesthesia_display() if op.anesthesia else '—'
                complication = op.complication or '—'
                op_lines.append(
                    f"{i}. {date} | {name} | Narkoz: {anesthesia} | Asorat: {complication}"
                )
            operations_text = '\n'.join(op_lines)
        else:
            operations_text = '—'

        row_data = [
            row_num - 1,
            patient.medical_record_number,
            patient.full_name,
            patient.get_gender_display(),
            patient.birth_date.strftime('%d.%m.%Y') if patient.birth_date else '',
            patient.get_resident_status_display(),
            get_full_address(patient),
            patient.get_social_status_display(),
            patient.passport_serial or '',
            patient.get_referral_type_display(),
            str(patient.referral_organization) if patient.referral_organization else '',
            patient.admission_diagnosis or '',
            patient.get_hours_after_illness_display(),
            'Ha' if patient.is_emergency else "Yo'q",
            'Ha' if patient.is_paid else "Yo'q",
            patient.admission_date.strftime('%d.%m.%Y %H:%M') if patient.admission_date else '',
            str(patient.department) if patient.department else '',
            patient.get_admission_count_display(),
            patient.days_in_hospital,
            patient.get_outcome_display(),
            str(patient.discharge_conclusion) if patient.discharge_conclusion else '',
            patient.discharge_date.strftime('%d.%m.%Y %H:%M') if patient.discharge_date else '',
            patient.clinical_main_diagnosis or '',
            str(patient.attending_doctor) if patient.attending_doctor else '',
            operations_text,
        ]

        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            cell.border = border

            # Yakun bo'yicha rang
            if col_num == 20:
                if value == 'Chiqarildi':
                    cell.fill = PatternFill("solid", fgColor="C6EFCE")
                elif value == 'Vafot etdi':
                    cell.fill = PatternFill("solid", fgColor="FFC7CE")
                elif value == "Boshqa shifoxonaga o'tkazildi":
                    cell.fill = PatternFill("solid", fgColor="FFEB9C")

    # Ustun kengliklari
    col_widths = [
        4, 14, 25, 8, 13, 10, 30, 15, 12, 16,
        22, 25, 18, 11, 8, 18, 18, 14, 12, 14,
        18, 18, 16, 22, 40
    ]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Jarrohlik ustuni qator balandligi
    for row_num in range(2, ws.max_row + 1):
        ops_cell = ws.cell(row=row_num, column=25)
        if ops_cell.value and '\n' in str(ops_cell.value):
            line_count = str(ops_cell.value).count('\n') + 1
            ws.row_dimensions[row_num].height = max(15, line_count * 15)

    # ==================== 2-SAHIFA: STATISTIKA ====================
    ws2 = wb.create_sheet("Statistika")

    ws2['A1'] = "Umumiy statistika"
    ws2['A1'].font = Font(bold=True, size=14)

    stats = [
        ("Jami bemorlar", qs_stats.count()),
        ("Chiqarildi", qs_stats.filter(outcome='discharged').count()),
        ("Vafot etdi", qs_stats.filter(outcome='deceased').count()),
        ("O'tkazildi", qs_stats.filter(outcome='transferred').count()),
        ("Shoshilinch qabul", qs_stats.filter(is_emergency=True).count()),
        ("Pullik yotqizilgan", qs_stats.filter(is_paid=True).count()),
        ("Rezidentlar", qs_stats.filter(resident_status='resident').count()),
        ("Norezidentlar", qs_stats.filter(resident_status='non_resident').count()),
        ("O'rtacha yotish kunlari", round(
            qs_stats.aggregate(avg=Avg('days_in_hospital'))['avg'] or 0, 1
        )),
    ]

    for row_num, (label, value) in enumerate(stats, 3):
        label_cell = ws2.cell(row=row_num, column=1, value=label)
        label_cell.font = Font(bold=True)
        ws2.cell(row=row_num, column=2, value=value)

    ws2.column_dimensions['A'].width = 30
    ws2.column_dimensions['B'].width = 15

    # ==================== 3-SAHIFA: JARROHLIK AMALIYOTLARI ====================
    ws3 = wb.create_sheet("Jarrohlik amaliyotlari")

    op_headers = [
        "№", "Bemor", "Bayonnoma №", "Sana",
        "Amaliyot nomi", "Narkoz", "Asorati"
    ]
    for col_num, header in enumerate(op_headers, 1):
        cell = ws3.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    op_row = 2
    for patient in qs:
        for op in patient.operations.all():
            name = str(op.operation_type) if op.operation_type else op.operation_name or '—'
            op_data = [
                op_row - 1,
                patient.full_name,
                patient.medical_record_number,
                op.operation_date.strftime('%d.%m.%Y') if op.operation_date else '—',
                name,
                op.get_anesthesia_display() if op.anesthesia else '—',
                op.complication or '—',
            ]
            for col_num, value in enumerate(op_data, 1):
                cell = ws3.cell(row=op_row, column=col_num, value=value)
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                cell.border = border
            op_row += 1

    if op_row == 2:
        ws3.cell(row=2, column=1, value="Jarrohlik amaliyotlari yo'q")

    op_col_widths = [4, 25, 14, 12, 35, 15, 25]
    for i, width in enumerate(op_col_widths, 1):
        ws3.column_dimensions[get_column_letter(i)].width = width

    # ==================== 4-SAHIFA: BEMORLAR + XIZMATLAR ====================
    from apps.services.models import PatientService
    from django.db.models import Sum, Count

    ws4 = wb.create_sheet("Bemorlar xizmatlari")
    ws4.column_dimensions['A'].width = 6
    ws4.column_dimensions['B'].width = 28
    ws4.column_dimensions['C'].width = 16
    ws4.column_dimensions['D'].width = 30
    ws4.column_dimensions['E'].width = 22
    ws4.column_dimensions['F'].width = 18
    ws4.column_dimensions['G'].width = 18

    # Sarlavha
    ws4.merge_cells('A1:G1')
    c = ws4.cell(row=1, column=1, value="BEMORLAR VA XIZMATLAR HISOBOTI")
    c.fill = header_fill; c.font = header_font
    c.alignment = center; c.border = border
    ws4.row_dimensions[1].height = 28

    # Ustun sarlavhalari
    h4 = ['№', 'Bemor', 'Bayonnoma', "Bo'lim", 'Kategoriya', 'Xizmat turi', "Summa (so'm)"]
    for col, h in enumerate(h4, 1):
        c = ws4.cell(row=2, column=col, value=h)
        c.fill = header_fill; c.font = header_font
        c.alignment = center; c.border = border
    ws4.row_dimensions[2].height = 22

    r4 = 3
    patient_num = 0

    for patient in qs:
        patient_services = PatientService.objects.filter(
            patient_card=patient
        ).select_related('service__category')

        if not patient_services.exists():
            continue

        patient_num += 1

        # Kategoriya bo'yicha guruhlash
        cat_stats = patient_services.values(
            'service__category__name'
        ).annotate(
            count=Count('id'),
            total=Sum('price'),
        ).order_by('service__category__name')

        patient_total = patient_services.aggregate(t=Sum('price'))['t'] or 0
        cat_count = cat_stats.count()
        start_row = r4

        # Har bir kategoriya qatori
        for i, cat in enumerate(cat_stats):
            if i == 0:
                # Birinchi qatorda bemor ma'lumotlari
                ws4.cell(row=r4, column=1, value=patient_num).border = border
                ws4.cell(row=r4, column=2, value=patient.full_name).border = border
                ws4.cell(row=r4, column=3, value=patient.medical_record_number).border = border
                ws4.cell(row=r4, column=4,
                         value=str(patient.department) if patient.department else '—').border = border
                ws4.cell(row=r4, column=5,
                         value=patient.get_patient_category_display()).border = border
            else:
                for col in range(1, 6):
                    ws4.cell(row=r4, column=col, value='').border = border

            # Kategoriya nomi
            c = ws4.cell(row=r4, column=6, value=cat['service__category__name'])
            c.fill = PatternFill('solid', fgColor='EBF5FB')
            c.font = Font(size=10); c.border = border

            # Kategoriya summa
            c7 = ws4.cell(row=r4, column=7, value=float(cat['total'] or 0))
            c7.fill = PatternFill('solid', fgColor='EBF5FB')
            c7.font = Font(size=10)
            c7.number_format = '#,##0'
            c7.alignment = Alignment(horizontal='right', vertical='center')
            c7.border = border

            ws4.row_dimensions[r4].height = 18
            r4 += 1

        # Bemorning 1-5 ustunlarini birlashtirish
        if cat_count > 1:
            for col in range(1, 6):
                ws4.merge_cells(
                    start_row=start_row, start_column=col,
                    end_row=r4 - 1, end_column=col
                )
                c = ws4.cell(row=start_row, column=col)
                c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Bemor jami
        ws4.merge_cells(start_row=r4, start_column=1, end_row=r4, end_column=6)
        c = ws4.cell(row=r4, column=1,
                     value=f"   {patient.full_name} — jami:")
        c.fill = PatternFill('solid', fgColor='D6E4F0')
        c.font = Font(bold=True, size=10)
        c.border = border

        c7 = ws4.cell(row=r4, column=7, value=float(patient_total))
        c7.fill = PatternFill('solid', fgColor='D6E4F0')
        c7.font = Font(bold=True, size=10)
        c7.number_format = '#,##0'
        c7.alignment = Alignment(horizontal='right', vertical='center')
        c7.border = border
        ws4.row_dimensions[r4].height = 20
        r4 += 1

    # Umumiy jami
    all_total = PatientService.objects.filter(
        patient_card__in=qs
    ).aggregate(t=Sum('price'))['t'] or 0

    ws4.merge_cells(start_row=r4, start_column=1, end_row=r4, end_column=6)
    c = ws4.cell(row=r4, column=1, value="UMUMIY JAMI:")
    c.fill = header_fill; c.font = header_font; c.border = border
    c7 = ws4.cell(row=r4, column=7, value=float(all_total))
    c7.fill = header_fill; c7.font = header_font
    c7.number_format = '#,##0'
    c7.alignment = Alignment(horizontal='right', vertical='center')
    c7.border = border
    ws4.row_dimensions[r4].height = 25

    if patient_num == 0:
        ws4.merge_cells('A3:G3')
        ws4.cell(row=3, column=1, value="Xizmatlar mavjud emas").border = border


    # ==================== 5-SAHIFA: BIRLASHGAN JADVAL ====================
    from apps.services.models import PatientService, ServiceCategory
    from django.db.models import Sum

    ws5 = wb.create_sheet("Bemor + Xizmatlar")

    # Barcha kategoriyalarni olish
    all_cats = list(ServiceCategory.objects.filter(is_active=True).order_by('name'))
    cat_names = [c.name for c in all_cats]

    # Ustun kengliklari
    fixed_cols = ['№', 'Bemor', 'Bayonnoma', "Bo'lim", 'Kategoriya', 'Qabul sanasi']
    all_headers = fixed_cols + cat_names + ["JAMI (so'm)"]

    ws5.column_dimensions['A'].width = 5
    ws5.column_dimensions['B'].width = 28
    ws5.column_dimensions['C'].width = 16
    ws5.column_dimensions['D'].width = 20
    ws5.column_dimensions['E'].width = 16
    ws5.column_dimensions['F'].width = 14
    for i in range(len(cat_names)):
        ws5.column_dimensions[get_column_letter(7 + i)].width = 16
    ws5.column_dimensions[get_column_letter(7 + len(cat_names))].width = 18

    # 1-qator: sarlavha
    total_cols = len(all_headers)
    ws5.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c = ws5.cell(row=1, column=1, value="BEMORLAR VA XIZMATLAR — BIRLASHGAN JADVAL")
    c.fill = header_fill; c.font = header_font
    c.alignment = center; c.border = border
    ws5.row_dimensions[1].height = 28

    # 2-qator: ustun sarlavhalari
    for col, h in enumerate(all_headers, 1):
        c = ws5.cell(row=2, column=col, value=h)
        c.fill = header_fill; c.font = header_font
        c.alignment = center; c.border = border
    ws5.row_dimensions[2].height = 35

    # Ma'lumotlar
    r5 = 3
    col_grand_totals = {name: 0 for name in cat_names}
    overall_total = 0

    for num, patient in enumerate(qs, 1):
        # Bemor xizmatlari kategoriya bo'yicha
        svc_by_cat = {}
        patient_svcs = PatientService.objects.filter(
            patient_card=patient
        ).values('service__category__name').annotate(total=Sum('price'))

        for row in patient_svcs:
            svc_by_cat[row['service__category__name']] = float(row['total'] or 0)

        patient_total = sum(svc_by_cat.values())
        overall_total += patient_total

        # Qator ma'lumotlari
        row_data = [
            num,
            patient.full_name,
            patient.medical_record_number,
            str(patient.department) if patient.department else '—',
            patient.get_patient_category_display(),
            patient.admission_date.strftime('%d.%m.%Y') if patient.admission_date else '—',
        ]

        # Har bir kategoriya uchun narx
        for cat_name in cat_names:
            val = svc_by_cat.get(cat_name, 0)
            row_data.append(val if val else None)
            col_grand_totals[cat_name] = col_grand_totals.get(cat_name, 0) + (val or 0)

        row_data.append(patient_total if patient_total else None)

        # Qatorni yozish
        for col, val in enumerate(row_data, 1):
            c = ws5.cell(row=r5, column=col, value=val)
            c.alignment = Alignment(
                horizontal='right' if col > 6 else 'left',
                vertical='center', wrap_text=True
            )
            c.border = border
            if col > 6 and val:
                c.number_format = '#,##0'
                # Xizmat bor kataklar — yashil fon
                c.fill = PatternFill('solid', fgColor='E9F7EF')
            if col == total_cols and val:
                c.font = Font(bold=True, size=10)
                c.fill = PatternFill('solid', fgColor='D6E4F0')

        # Alternativ qator rangi
        if num % 2 == 0:
            for col in range(1, 7):
                c = ws5.cell(row=r5, column=col)
                if not c.fill or c.fill.fgColor.rgb == '00000000':
                    c.fill = PatternFill('solid', fgColor='F8F9FA')

        ws5.row_dimensions[r5].height = 18
        r5 += 1

    # Jami qator
    total_row_data = ['', 'JAMI:', '', '', '', '']
    for cat_name in cat_names:
        t = col_grand_totals.get(cat_name, 0)
        total_row_data.append(t if t else None)
    total_row_data.append(overall_total)

    for col, val in enumerate(total_row_data, 1):
        c = ws5.cell(row=r5, column=col, value=val)
        c.fill = header_fill; c.font = header_font
        c.border = border
        c.alignment = Alignment(
            horizontal='right' if col > 6 else 'left',
            vertical='center'
        )
        if col > 6 and val:
            c.number_format = '#,##0'
    ws5.row_dimensions[r5].height = 25

    if r5 == 3:
        ws5.merge_cells(start_row=3, start_column=1, end_row=3, end_column=total_cols)
        ws5.cell(row=3, column=1, value="Bemorlar topilmadi").border = border


    # ==================== 6-SAHIFA: TASHKILOT BO'YICHA ====================
    from apps.patients.models import Organization
    from apps.services.models import PatientService

    ws6 = wb.create_sheet("Tashkilotlar")
    ws6.column_dimensions['A'].width = 5
    ws6.column_dimensions['B'].width = 40
    ws6.column_dimensions['C'].width = 12
    ws6.column_dimensions['D'].width = 12
    ws6.column_dimensions['E'].width = 12
    ws6.column_dimensions['F'].width = 20
    ws6.column_dimensions['G'].width = 20
    ws6.column_dimensions['H'].width = 20

    # Sarlavha
    ws6.merge_cells('A1:H1')
    c = ws6.cell(row=1, column=1, value="TASHKILOT BO'YICHA STATISTIKA")
    c.fill = header_fill; c.font = header_font
    c.alignment = center; c.border = border
    ws6.row_dimensions[1].height = 28

    # Ustun sarlavhalar
    h6 = [
        '№', 'Tashkilot', 'Korxona kodi', 'Filial kodi',
        'Bemorlar', "Bo'lim bo'yicha", 'Tashxis', "Xizmatlar summasi (so'm)"
    ]
    for col, h in enumerate(h6, 1):
        c = ws6.cell(row=2, column=col, value=h)
        c.fill = header_fill; c.font = header_font
        c.alignment = center; c.border = border
    ws6.row_dimensions[2].height = 25

    # Tashkilotlar
    org_qs = (
        qs.filter(workplace_org__isnull=False)
        .values(
            'workplace_org__id',
            'workplace_org__enterprise_name',
            'workplace_org__branch_name',
            'workplace_org__enterprise_code',
            'workplace_org__branch_code',
        )
        .annotate(patient_count=Count('id'))
        .order_by('-patient_count')
    )

    r6 = 3
    grand_patients = 0
    grand_svc_total = 0

    for num, org in enumerate(org_qs, 1):
        org_id = org['workplace_org__id']
        ent_name = org['workplace_org__enterprise_name'] or ''
        branch   = org['workplace_org__branch_name'] or ''
        full_name = f"{ent_name} — {branch}" if branch else ent_name
        p_count  = org['patient_count']
        grand_patients += p_count

        # Xizmatlar summasi
        svc_total = PatientService.objects.filter(
            patient_card__workplace_org_id=org_id
        ).filter(
            patient_card__in=qs
        ).aggregate(t=Sum('price'))['t'] or 0
        svc_total = float(svc_total)
        grand_svc_total += svc_total

        # Bo'lim taqsimoti
        dept_dist = list(
            qs.filter(workplace_org_id=org_id)
            .values('department__name')
            .annotate(cnt=Count('id'))
            .order_by('-cnt')[:3]
        )
        dept_str = ', '.join([
            f"{d['department__name']}({d['cnt']})"
            for d in dept_dist if d['department__name']
        ]) or '—'

        # Tashxis (eng ko'p uchraydigan)
        diag_dist = list(
            qs.filter(workplace_org_id=org_id)
            .exclude(admission_diagnosis='')
            .values('admission_diagnosis')
            .annotate(cnt=Count('id'))
            .order_by('-cnt')[:1]
        )
        diag_str = diag_dist[0]['admission_diagnosis'][:50] if diag_dist else '—'

        row6 = [
            num, full_name,
            org['workplace_org__enterprise_code'] or '—',
            org['workplace_org__branch_code'] or '—',
            p_count, dept_str, diag_str, svc_total
        ]

        for col, val in enumerate(row6, 1):
            c = ws6.cell(row=r6, column=col, value=val)
            c.alignment = Alignment(
                horizontal='right' if col in (5, 8) else 'left',
                vertical='center', wrap_text=True
            )
            c.border = border
            if col == 8:
                c.number_format = '#,##0'
            if num % 2 == 0 and col < 6:
                c.fill = PatternFill('solid', fgColor='F8F9FA')

        ws6.row_dimensions[r6].height = 18
        r6 += 1

    # Jami
    ws6.merge_cells(start_row=r6, start_column=1, end_row=r6, end_column=4)
    c = ws6.cell(row=r6, column=1, value="JAMI:")
    c.fill = header_fill; c.font = header_font; c.border = border

    c5 = ws6.cell(row=r6, column=5, value=grand_patients)
    c5.fill = header_fill; c5.font = header_font
    c5.alignment = Alignment(horizontal='right', vertical='center')
    c5.border = border

    ws6.merge_cells(start_row=r6, start_column=6, end_row=r6, end_column=7)
    c67 = ws6.cell(row=r6, column=6, value='')
    c67.fill = header_fill; c67.border = border

    c8 = ws6.cell(row=r6, column=8, value=grand_svc_total)
    c8.fill = header_fill; c8.font = header_font
    c8.number_format = '#,##0'
    c8.alignment = Alignment(horizontal='right', vertical='center')
    c8.border = border
    ws6.row_dimensions[r6].height = 25

    if r6 == 3:
        ws6.merge_cells('A3:H3')
        ws6.cell(row=3, column=1,
                 value="Tashkilotga biriktirilgan bemorlar yo'q").border = border

    # ==================== DORI SHEETI ====================
    from apps.services.models import PatientMedicine
    from django.db.models import Sum as MSum

    ws_med = wb.create_sheet("Dori-darmonlar")
    ws_med.column_dimensions['A'].width = 5
    ws_med.column_dimensions['B'].width = 40
    ws_med.column_dimensions['C'].width = 12
    ws_med.column_dimensions['D'].width = 12
    ws_med.column_dimensions['E'].width = 15
    ws_med.column_dimensions['F'].width = 20

    ws_med.merge_cells('A1:F1')
    c = ws_med.cell(row=1, column=1, value="DORI-DARMON STATISTIKASI")
    c.fill = header_fill; c.font = header_font
    c.alignment = center; c.border = border
    ws_med.row_dimensions[1].height = 28

    heads = ['№', 'Dori nomi', 'Birlik', 'Jami miqdor', 'Bemorlar soni', "Jami summa (so'm)"]
    for col, h in enumerate(heads, 1):
        c = ws_med.cell(row=2, column=col, value=h)
        c.fill = header_fill; c.font = header_font
        c.alignment = center; c.border = border
    ws_med.row_dimensions[2].height = 24

    top_meds = (
        PatientMedicine.objects
        .filter(patient_card__in=qs)
        .values('medicine__name', 'medicine__unit')
        .annotate(total_qty=MSum('quantity'), total_sum=MSum('price'), cnt=Count('patient_card', distinct=True))
        .order_by('-total_sum')
    )

    med_grand = 0
    for ri, m in enumerate(top_meds, 1):
        row_data = [
            ri,
            m['medicine__name'],
            m['medicine__unit'],
            float(m['total_qty'] or 0),
            m['cnt'],
            float(m['total_sum'] or 0),
        ]
        med_grand += float(m['total_sum'] or 0)
        for col, val in enumerate(row_data, 1):
            c = ws_med.cell(row=ri+2, column=col, value=val)
            c.alignment = Alignment(horizontal='center' if col in (1,3,4,5) else ('right' if col==6 else 'left'), vertical='center')
            c.border = border
            if col == 6: c.number_format = '#,##0'
            if ri % 2 == 0: c.fill = PatternFill('solid', fgColor='F8F9FA')
        ws_med.row_dimensions[ri+2].height = 18

    # Jami qator
    last_r = len(list(top_meds)) + 3
    ws_med.merge_cells(start_row=last_r, start_column=1, end_row=last_r, end_column=5)
    c = ws_med.cell(row=last_r, column=1, value="JAMI:")
    c.fill = header_fill; c.font = header_font; c.border = border
    c6 = ws_med.cell(row=last_r, column=6, value=med_grand)
    c6.fill = header_fill; c6.font = header_font
    c6.number_format = '#,##0'
    c6.alignment = Alignment(horizontal='right', vertical='center')
    c6.border = border
    ws_med.row_dimensions[last_r].height = 24

    # ==================== JAVOB ====================
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename="bemorlar_royxati.xlsx"'
    wb.save(response)
    return response


def export_pdf(request):
    qs = get_filtered_queryset(request)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        rightMargin=1 * cm, leftMargin=1 * cm,
        topMargin=1.5 * cm, bottomMargin=1 * cm
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'title', parent=styles['Heading1'],
        fontSize=14, alignment=1, spaceAfter=12
    )
    small_style = ParagraphStyle(
        'small', parent=styles['Normal'],
        fontSize=7, leading=9
    )

    elements = []
    elements.append(Paragraph(
        "Shifoxonadan chiqarilganlar statistik ro'yxati", title_style
    ))
    elements.append(Spacer(1, 0.3 * cm))

    table_headers = [
        "№", "Bayonnoma", "Ism-familiya", "Jins", "Tug'ilgan\nsana",
        "Bo'lim", "Yotqizilgan\nsana", "Yotgan\nkun",
        "Yakun", "Xulosa", "MKB-10", "Shifokor"
    ]

    table_data = [[Paragraph(h, small_style) for h in table_headers]]

    for i, patient in enumerate(qs, 1):
        row = [
            str(i),
            patient.medical_record_number or '',
            Paragraph(patient.full_name or '', small_style),
            patient.get_gender_display(),
            patient.birth_date.strftime('%d.%m.%Y') if patient.birth_date else '',
            Paragraph(str(patient.department) if patient.department else '', small_style),
            patient.admission_date.strftime('%d.%m.%Y') if patient.admission_date else '',
            str(patient.days_in_hospital),
            patient.get_outcome_display(),
            str(patient.discharge_conclusion) if patient.discharge_conclusion else '—',
            patient.clinical_main_diagnosis or '',
            Paragraph(
                str(patient.attending_doctor) if patient.attending_doctor else '', small_style
            ),
        ]
        table_data.append(row)

    col_widths = [
        1 * cm, 2.5 * cm, 4.5 * cm, 1.2 * cm, 2.2 * cm,
        3.5 * cm, 2.5 * cm, 1.5 * cm, 2.5 * cm, 2.5 * cm, 1.8 * cm, 3.8 * cm
    ]

    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [
            colors.white, colors.HexColor('#F2F2F2')
        ]),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)

    response = HttpResponse(buffer, content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="bemorlar_royxati.pdf"'
    return response


# apps/statistic/exports.py — get_filtered_queryset yangilash
