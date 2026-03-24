# apps/statistic/urls.py — TO'G'RI
from django.urls import path
from . import views
from .exports import export_excel, export_pdf

urlpatterns = [
    path('', views.statistics_dashboard, name='statistics_dashboard'),
    path('export/excel/', export_excel, name='export_excel'),
    path('export/pdf/', export_pdf, name='export_pdf'),
]