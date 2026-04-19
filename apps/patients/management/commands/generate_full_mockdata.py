# apps/patients/management/commands/generate_full_mockdata.py

import random
import uuid
from decimal import Decimal
from datetime import datetime, timedelta

from django.core.management.base import BaseCommand
from django.utils import timezone

from apps.patients.models import (
    PatientCard, Organization, Department, Doctor,
    Country, Region, District, City, DeathCause,
    SurgicalOperation, OperationType,
)
from apps.services.models import (
    Service, ServiceCategory, PatientService,
    Medicine, PatientMedicine,
)


ERKAK = [
    "Karimov Bobur", "Rahimov Jasur", "Toshmatov Ulugbek",
    "Hasanov Sherzod", "Yusupov Akbar", "Nazarov Dilshod",
    "Qodirov Mansur", "Mirzayev Sardor", "Aliyev Firdavs",
    "Ergashev Nodir", "Sobirov Eldor", "Xoliqov Anvar",
    "Abdullayev Sanjar", "Normatov Zafar", "Usmonov Ibrohim",
    "Tursunov Hamid", "Baxtiyorov Otabek", "Holmatov Bekzod",
    "Sultonov Jahongir", "Nishonov Ravshan", "Jurayev Timur",
    "Ismoilov Laziz", "Raximov Murod", "Xasanov Lochinbek",
]
AYOL = [
    "Karimova Nilufar", "Rahimova Malika", "Toshmatova Dilorom",
    "Hasanova Gulnora", "Yusupova Feruza", "Nazarova Shahnoza",
    "Qodirova Zulfiya", "Mirzayeva Mohira", "Aliyeva Maftuna",
    "Ergasheva Nasiba", "Sobirova Kamola", "Xoliqova Nargiza",
    "Abdullayeva Barno", "Normatova Hulkar", "Usmonova Sabohat",
    "Tursunova Gavhar", "Baxtiyorova Munira", "Holmatova Sevinch",
    "Sultonova Iroda", "Nishonova Zulfiya", "Jurayeva Dilnoza",
    "Ismoilova Sarvinoz", "Raximova Lobar", "Xasanova Oydin",
]
MKB = [
    ("I21.0", "O'tkir miokard infarkti"),
    ("I10",   "Gipertoniya kasalligi"),
    ("J18.9", "Pnevmoniya"),
    ("K35.8", "O'tkir appenditsit"),
    ("S72.0", "Son suyagi sinishi"),
    ("I63.9", "Insult"),
    ("K80.1", "Xoletsistit"),
    ("N20.0", "Buyrak toshi"),
    ("J44.1", "SOOB"),
    ("E11.9", "Qandli diabet 2-tur"),
    ("I50.0", "Yurak yetishmovchiligi"),
    ("K29.7", "Gastrit"),
    ("M16.1", "Koksartroz"),
    ("G35",   "Ko'p skleroz"),
    ("N10",   "O'tkir pielonefrit"),
    ("K85.9", "O'tkir pankreatit"),
    ("I20.0", "Stenokardiya"),
    ("J45.9", "Bronxial astma"),
]
KOCHALAR = [
    "Amir Temur ko'chasi", "Mustaqillik ko'chasi",
    "Navoiy ko'chasi", "Olmazor ko'chasi",
    "Bog'ishamol ko'chasi", "Chilonzor ko'chasi",
]
OPERATIONS = [
    "Appendektomiya", "Xoletsistektomiya", "Koronar shuntlash",
    "Artroplastika", "Nefrektomiya", "Grija plastikasi",
    "Gastroektomiya", "Rezektsiya", "Laparoskopiya",
    "Traxeotomiya", "Amputatsiya", "Bypass operatsiyasi",
]
SERVICES_DATA = [
    ("Laboratoriya", "lab", [
        ("Umumiy qon tahlili", 90000, 46930),
        ("Umumiy siydik tahlili", 70000, 29966),
        ("Glyukoza tahlili", 65000, 32786),
        ("Bilirubin", 50000, 53478),
        ("Ferment ALT", 50000, 36579),
        ("Gemoglobin tahlili", 30000, 17398),
        ("Leykotsitlar tahlili", 30000, 17398),
        ("Koagulogramma", 115000, 65539),
    ]),
    ("Rentgen", "radiology", [
        ("Ko'krak rentgeni", 80000, 45000),
        ("Qorin bo'shlig'i rentgeni", 90000, 50000),
        ("Bosh suyagi rentgeni", 85000, 48000),
        ("Son suyagi rentgeni", 95000, 55000),
    ]),
    ("UZI", "radiology", [
        ("Qorin bo'shlig'i UZI", 120000, 68000),
        ("Buyrak UZI", 110000, 62000),
        ("Yurak EXO-KG", 180000, 95000),
        ("Qalqonsimon bez UZI", 100000, 58000),
    ]),
    ("Shifokor ko'rigi", "consultation", [
        ("Kardiolog ko'rigi", 80000, 45000),
        ("Nevropatolog ko'rigi", 75000, 42000),
        ("Endokrinolog ko'rigi", 80000, 45000),
        ("Urolog ko'rigi", 70000, 40000),
        ("Ginekolog ko'rigi", 75000, 42000),
    ]),
    ("Fizioterapiya", "physio", [
        ("Elektroforez", 45000, 25000),
        ("Massaj (1 seans)", 60000, 35000),
        ("Magnit terapiya", 50000, 28000),
        ("UFO terapiya", 40000, 22000),
    ]),
    ("Jarrohlik xizmatlari", "surgery", [
        ("Operatsiya xonasi (kichik)", 350000, 200000),
        ("Operatsiya xonasi (o'rta)", 600000, 350000),
        ("Operatsiya xonasi (katta)", 1200000, 700000),
        ("Narkoz (mahalliy)", 150000, 85000),
        ("Narkoz (umumiy)", 500000, 290000),
    ]),
    ("Yotoq-joy", "other", [
        ("Kunlik yotoq-joy (oddiy)", 180000, 105000),
        ("Kunlik yotoq-joy (2 kishilik)", 250000, 145000),
        ("Kunlik yotoq-joy (lyuks)", 550000, 324000),
    ]),
]
MEDICINES_DATA = [
    ("Sefiksim 400mg", "kapsula", 25000),
    ("Amoksisiklin 500mg", "kapsula", 8000),
    ("Metronidazol 500mg", "tabletka", 5000),
    ("Omeprazol 20mg", "kapsula", 6000),
    ("Diklofenak 75mg/3ml", "ampula", 7000),
    ("Ketorol 30mg/ml", "ampula", 9000),
    ("NaCl 0.9% 500ml", "shisha", 12000),
    ("Glukoza 5% 500ml", "shisha", 14000),
    ("Heparin 5000 IU/ml", "ampula", 35000),
    ("Furosemid 40mg", "tabletka", 3000),
    ("Enalapril 10mg", "tabletka", 4000),
    ("Metformin 1000mg", "tabletka", 5500),
    ("Prednizolon 30mg", "ampula", 8500),
    ("Aktovegin 200mg", "ampula", 28000),
    ("Mexiletil 100mg", "kapsula", 15000),
    ("Vitamin C 500mg", "tabletka", 2500),
    ("B12 vitamini 500mcg", "ampula", 6500),
    ("Albumin 20% 100ml", "shisha", 180000),
    ("Pantoprazol 40mg", "tabletka", 7000),
    ("Amlodipin 10mg", "tabletka", 4500),
]


class Command(BaseCommand):
    help = "To'liq test ma'lumotlari: bemorlar + xizmatlar + dorilar + operatsiyalar"

    def add_arguments(self, parser):
        parser.add_argument('--count', type=int, default=30,
                            help='Nechta bemor (default: 30)')
        parser.add_argument('--clear', action='store_true',
                            help='Avval bemorlar va xizmatlarni tozalash')

    def handle(self, *args, **options):
        count  = options['count']

        if options['clear']:
            PatientService.objects.all().delete()
            PatientMedicine.objects.all().delete()
            SurgicalOperation.objects.all().delete()
            PatientCard.objects.all().delete()
            self.stdout.write("Eski ma'lumotlar o'chirildi.")

        # ===== LOOKUP DATA =====
        country, _ = Country.objects.get_or_create(name="O'zbekiston")
        region,  _ = Region.objects.get_or_create(name="Toshkent shahri", country=country)

        dist_names = ["Chilonzor", "Yunusobod", "Mirzo Ulug'bek", "Shayxontohur", "Yakkasaroy"]
        districts  = []
        for n in dist_names:
            d, _ = District.objects.get_or_create(name=n, region=region)
            districts.append(d)

        cities = []
        for d in districts:
            c, _ = City.objects.get_or_create(name=f"{d.name} MFY", district=d)
            cities.append(c)

        dept_names = [
            "Terapiya", "Jarrohlik", "Kardiologiya", "Nevrologiya",
            "Ortopediya", "Ginekologiya", "Pediatriya", "Reanimatsiya",
            "Travmatologiya", "Endokrinologiya",
        ]
        depts = []
        for n in dept_names:
            d, _ = Department.objects.get_or_create(name=n)
            depts.append(d)

        doc_names = [
            ("Karimov Bobur Aliyevich", False),
            ("Rahimova Nilufar Hasanovna", False),
            ("Toshmatov Jasur Mirzayevich", False),
            ("Yusupova Malika Karimovna", False),
            ("Hasanov Sherzod Umarovich", True),
            ("Nazarova Dilorom Ergashevna", True),
        ]
        doctors = []
        for full_name, is_head in doc_names:
            d, _ = Doctor.objects.get_or_create(
                full_name=full_name,
                defaults={'department': random.choice(depts), 'is_head': is_head}
            )
            doctors.append(d)

        # Temir yo'l tashkilotlari
        ty_orgs = []
        ty_names = [
            "O'zbekiston temir yo'llari bosh idorasi",
            "Toshkent temir yo'l bo'limi",
            "Samarqand temir yo'l bo'limi",
            "Andijon temir yo'l bo'limi",
            "Buxoro temir yo'l bo'limi",
        ]
        for n in ty_names:
            o, _ = Organization.objects.get_or_create(enterprise_name=n, defaults={'branch_name': '', 'is_active': True})
            ty_orgs.append(o)

        # ===== XIZMATLAR VA DORILAR =====
        services_by_cat = {}
        for cat_name, cat_type, svc_list in SERVICES_DATA:
            cat, _ = ServiceCategory.objects.get_or_create(
                name=cat_name,
                defaults={'category_type': cat_type, 'is_active': True}
            )
            svcs = []
            for svc_name, price_n, price_r in svc_list:
                svc, _ = Service.objects.get_or_create(
                    name=svc_name, category=cat,
                    defaults={
                        'price_normal': price_n,
                        'price_railway': price_r,
                        'is_active': True,
                    }
                )
                svcs.append(svc)
            services_by_cat[cat_name] = svcs

        all_services = [s for svcs in services_by_cat.values() for s in svcs]

        medicines = []
        for med_name, unit, price in MEDICINES_DATA:
            m, _ = Medicine.objects.get_or_create(
                name=med_name,
                defaults={'unit': unit, 'is_active': True}
            )
            medicines.append((m, price))

        op_types = []
        for op_name in OPERATIONS:
            ot, _ = OperationType.objects.get_or_create(name=op_name)
            op_types.append(ot)

        # ===== BEMORLAR YARATISH =====
        # Har xil kategoriya: TY, pullik, norezident, ambulator
        categories = (
            ['railway'] * 14 +
            ['paid']    * 8  +
            ['non_resident'] * 4 +
            ['railway'] * 4   # ambulator TY
        )
        random.shuffle(categories[:26])

        created = 0
        for i in range(count):
            category    = categories[i % len(categories)]
            is_ambulatory = (i % 7 == 0)  # har 7 chi bemor ambulator
            gender      = random.choice(['M', 'F'])
            full_name   = random.choice(ERKAK if gender == 'M' else AYOL)
            mkb_code, mkb_text = random.choice(MKB)
            dept        = random.choice(depts)
            doctor      = random.choice(doctors)
            district    = random.choice(districts)
            city        = City.objects.filter(district=district).first()

            # Sana
            days_ago    = random.randint(1, 365)
            admission   = timezone.now() - timedelta(days=days_ago)
            days_in     = 0 if is_ambulatory else random.randint(3, 25)
            discharge   = admission + timedelta(days=days_in) if not is_ambulatory else None

            outcome     = random.choices(
                ['discharged', 'deceased', 'transferred'],
                weights=[70, 10, 20]
            )[0]
            status      = 'completed'

            resident    = 'non_resident' if category == 'non_resident' else 'resident'
            passport    = f"AB{random.randint(1000000,9999999)}" if resident == 'resident' else 'FOREIGN'

            # Ish joyi (TY uchun)
            workplace_org = random.choice(ty_orgs) if category == 'railway' else None

            # Yil
            birth_y = random.randint(1950, 2005)
            if is_ambulatory and random.random() < 0.2:
                birth_y = random.randint(2010, 2020)  # bola

            while True:
                rec_num = f"{'AMB' if is_ambulatory else 'STA'}-{datetime.now().year}-{str(uuid.uuid4())[:6].upper()}"
                if not PatientCard.objects.filter(medical_record_number=rec_num).exists():
                    break

            try:
                patient = PatientCard.objects.create(
                    medical_record_number = rec_num,
                    full_name             = full_name,
                    gender                = gender,
                    birth_date            = datetime(birth_y, random.randint(1,12), random.randint(1,28)).date(),
                    resident_status       = resident,
                    passport_serial       = passport,
                    country               = country,
                    region                = region,
                    district              = district,
                    city                  = city,
                    street_address        = f"{random.choice(KOCHALAR)}, {random.randint(1,120)}-uy",
                    social_status         = random.choice(['employed','unemployed','pensioner','student_higher']),
                    patient_category      = category,
                    referral_type         = random.choice(['self','ambulance','referral','liniya']),
                    admission_diagnosis   = f"{mkb_code} - {mkb_text}",
                    clinical_main_diagnosis = mkb_code,
                    clinical_main_diagnosis_text = mkb_text,
                    is_emergency          = random.choice([True, False]),
                    is_paid               = (category != 'railway'),
                    admission_date        = admission,
                    department            = None if is_ambulatory else dept,
                    attending_doctor      = doctor,
                    days_in_hospital      = days_in,
                    discharge_date        = discharge,
                    outcome               = outcome,
                    status                = status,
                    visit_type            = 'ambulatory' if is_ambulatory else 'inpatient',
                    workplace_org         = workplace_org,
                    position              = "Lokomotiv brigadasi boshlig'i" if category == 'railway' else '',
                )

                # ===== XIZMATLAR =====
                n_services = random.randint(1, 5)
                chosen_svcs = random.sample(all_services, min(n_services, len(all_services)))
                for svc in chosen_svcs:
                    qty   = random.randint(1, 3)
                    price = svc.price_railway if category == 'railway' else svc.price_normal
                    if category == 'non_resident':
                        price = round(svc.price_normal * Decimal('1.25'), 0)
                    PatientService.objects.create(
                        patient_card              = patient,
                        service                   = svc,
                        quantity                  = qty,
                        price                     = price,
                        status                    = random.choice(['ordered','completed','completed']),
                        patient_category_at_order = category,
                        ordered_by                = doctor,
                        ordered_at                = admission + timedelta(hours=random.randint(1,48)),
                        notes                     = '',
                    )

                # ===== DORILAR =====
                if not is_ambulatory:
                    n_meds = random.randint(2, 6)
                    chosen_meds = random.sample(medicines, min(n_meds, len(medicines)))
                    for med, base_price in chosen_meds:
                        qty   = Decimal(str(random.randint(1, 10)))
                        price = Decimal(str(base_price))
                        if category == 'railway':
                            price = round(price * Decimal('0.7'), 0)
                        elif category == 'non_resident':
                            price = round(price * Decimal('1.3'), 0)
                        PatientMedicine.objects.create(
                            patient_card = patient,
                            medicine     = med,
                            quantity     = qty,
                            price        = price,
                            ordered_by   = doctor,
                            ordered_at   = admission + timedelta(hours=random.randint(2,72)),
                        )

                # ===== OPERATSIYALAR =====
                if not is_ambulatory and random.random() < 0.35:
                    n_ops = random.randint(1, 2)
                    for _ in range(n_ops):
                        op_type = random.choice(op_types)
                        SurgicalOperation.objects.create(
                            patient_card   = patient,
                            operation_type = op_type,
                            operation_date = (admission + timedelta(days=random.randint(0,3))).date(),
                            anesthesia     = random.choice(['yes','local','no']),
                            complication   = random.choice(['','','','Qon ketish','Infeksiya']),
                        )

                # ===== VAFOT SABABI =====
                if outcome == 'deceased':
                    DeathCause.objects.create(
                        patient_card    = patient,
                        immediate_cause = f"{mkb_text} asoratlari",
                        underlying_cause= mkb_text,
                        main_disease_code = mkb_code,
                    )

                created += 1
                if created % 5 == 0:
                    self.stdout.write(f"  {created} ta yaratildi...")

            except Exception as e:
                self.stdout.write(self.style.WARNING(f"  Xato ({i}): {e}"))

        self.stdout.write(self.style.SUCCESS(
            f"\n✅ {created} ta bemor yaratildi!"
            f"\n   TY bemorlar: ~{int(created*0.6)} ta"
            f"\n   Pullik: ~{int(created*0.27)} ta"
            f"\n   Norezident: ~{int(created*0.13)} ta"
            f"\n   Ambulatorlar: ~{created//7} ta"
            f"\n   Xizmatlar: {PatientService.objects.count()} ta"
            f"\n   Dorilar: {PatientMedicine.objects.count()} ta"
            f"\n   Operatsiyalar: {SurgicalOperation.objects.count()} ta"
        ))