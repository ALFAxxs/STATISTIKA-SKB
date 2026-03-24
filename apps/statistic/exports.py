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

    return qs
