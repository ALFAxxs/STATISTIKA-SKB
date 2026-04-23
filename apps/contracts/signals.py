# apps/contracts/signals.py

from django.db.models.signals import post_save
from django.dispatch import receiver
from django.core.files.base import ContentFile
from apps.patients.models import PatientCard


@receiver(post_save, sender=PatientCard)
def create_contract_on_admission(sender, instance, created, **kwargs):
    """Bemor yaratilganda avtomatik shartnoma shakllantiradi"""
    if not created:
        return

    # Faqat pullik va norezident statsionar bemorlar uchun
    if instance.patient_category not in ('paid', 'non_resident'):
        return
    if instance.visit_type == 'ambulatory':
        return

    from apps.contracts.models import Contract
    from apps.contracts.utils import generate_contract_pdf

    # Shartnoma raqami yaratish
    year = instance.admission_date.year if instance.admission_date else 2026
    count = Contract.objects.filter(created_at__year=year).count() + 1
    contract_number = f"SHA-{year}-{count:04d}"

    # Shartnoma yaratish
    contract = Contract.objects.create(
        patient_card=instance,
        contract_number=contract_number,
        contract_type=instance.patient_category,
        status='active',
    )

    # PDF generatsiya va saqlash
    try:
        pdf_bytes = generate_contract_pdf(contract)
        filename = f"contract_{contract_number}.pdf"
        contract.pdf_file.save(filename, ContentFile(pdf_bytes), save=True)
    except Exception as e:
        # PDF generatsiya xatosi bo'lsa ham shartnoma yozuvi saqlanadi
        import logging
        logging.getLogger(__name__).error(f"Contract PDF error: {e}")