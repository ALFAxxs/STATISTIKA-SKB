# apps/statistic/views.py

from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from django.db.models import Count, Avg, Q
from django.db.models.functions import TruncMonth
from apps.patients.models import PatientCard, Department
import json


@login_required
def statistics_dashboard(request):

    # --- Filtrlar ---
    year = request.GET.get('year', '')
    department_id = request.GET.get('department', '')

    qs = PatientCard.objects.all()
    if year:
        qs = qs.filter(admission_date__year=year)
    if department_id:
        qs = qs.filter(department_id=department_id)

    # --- Umumiy sonlar ---
    total = qs.count()
    discharged = qs.filter(outcome='discharged').count()
    deceased = qs.filter(outcome='deceased').count()
    transferred = qs.filter(outcome='transferred').count()

    # --- Jins bo'yicha ---
    gender_stats = qs.values('gender').annotate(count=Count('id'))
    gender_data = {'M': 0, 'F': 0}
    for item in gender_stats:
        if item['gender'] in gender_data:
            gender_data[item['gender']] = item['count']

    # --- Bo'lim bo'yicha ---
    dept_stats = (
        qs.values('department__name')
        .annotate(count=Count('id'))
        .order_by('-count')
    )
    # None bo'lgan bo'limlarni filtrlaymiz
    dept_stats = [
        item for item in dept_stats
        if item['department__name']
    ]

    # --- Oylik dinamika ---
    monthly_stats = (
        qs.annotate(month=TruncMonth('admission_date'))
        .values('month')
        .annotate(count=Count('id'))
        .order_by('month')
    )
    monthly_labels = [
        item['month'].strftime('%Y-%m')
        for item in monthly_stats
        if item['month']
    ]
    monthly_values = [
        item['count']
        for item in monthly_stats
        if item['month']
    ]

    # --- Ijtimoiy holat ---
    social_stats = (
        qs.values('social_status')
        .annotate(count=Count('id'))
        .order_by('-count')
    )

    # --- Rezident / Norezident ---
    resident_count = qs.filter(resident_status='resident').count()
    non_resident_count = qs.filter(resident_status='non_resident').count()

    # --- Shoshilinch vs oddiy ---
    emergency_count = qs.filter(is_emergency=True).count()
    non_emergency_count = qs.filter(is_emergency=False).count()

    # --- O'rtacha yotish kunlari ---
    avg_days = qs.aggregate(avg=Avg('days_in_hospital'))['avg'] or 0

    # --- Yillar ro'yxati (filter uchun) ---  # ← tuzatildi
    years = (
        PatientCard.objects
        .exclude(admission_date=None)
        .dates('admission_date', 'year', order='DESC')
    )
    year_list = [d.year for d in years]

    departments = Department.objects.filter(is_active=True)

    return render(request, 'statistic/dashboard.html', {
        'total': total,
        'discharged': discharged,
        'deceased': deceased,
        'transferred': transferred,
        'gender_data': json.dumps(gender_data),
        'dept_stats': dept_stats,
        'monthly_labels': json.dumps(monthly_labels),
        'monthly_values': json.dumps(monthly_values),
        'social_stats': social_stats,
        'resident_count': resident_count,
        'non_resident_count': non_resident_count,
        'emergency_count': emergency_count,
        'non_emergency_count': non_emergency_count,
        'avg_days': round(avg_days, 1),
        'years': year_list,                          # ← tuzatildi
        'departments': departments,
        'selected_year': year,
        'selected_dept': department_id,
    })