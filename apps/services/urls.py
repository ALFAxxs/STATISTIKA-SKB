# apps/services/urls.py

from django.urls import path
from . import views

urlpatterns = [
    # Bemor xizmatlari
    path('patient/<int:patient_pk>/', views.patient_services, name='patient_services'),
    path('patient/<int:patient_pk>/add/', views.add_service, name='add_service'),
    path('<int:pk>/update/', views.update_service, name='update_service'),
    path('<int:pk>/delete/', views.delete_service, name='delete_service'),

    # AJAX
    path('search/', views.service_search, name='service_search'),

    # Statistika
    path('statistics/', views.service_statistics, name='service_statistics'),

    # Export
    path('export/excel/', views.export_services_excel, name='export_services_excel'),
    path('export/pdf/', views.export_services_pdf, name='export_services_pdf'),
]
