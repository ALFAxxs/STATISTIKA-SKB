# skb_stat/urls.py

from django.contrib import admin
from django.urls import path, include
from django.shortcuts import redirect

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', lambda request: redirect('patient_list'), name='home'),  # ← qo'shish
    path('patients/', include('apps.patients.urls')),
    path('statistics/', include('apps.statistic.urls')),
    path('login/', include('apps.users.urls')),
    path('logout/', include('apps.users.urls')),
    path('users/', include('apps.users.urls')),
    path('create/', include('apps.users.urls')),
    path('access-denied/', include('apps.users.urls')),
]