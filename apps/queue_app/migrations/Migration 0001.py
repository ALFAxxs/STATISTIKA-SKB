from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
        ('patients', '0001_initial'),
        ('services', '0001_initial'),
    ]

    operations = [
        migrations.CreateModel(
            name='QueueTicket',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True)),
                ('ticket_number', models.PositiveIntegerField(verbose_name='Navbat raqami')),
                ('status', models.CharField(
                    choices=[
                        ('waiting',  'Kutmoqda'),
                        ('calling',  "Chaqirilmoqda"),
                        ('serving',  "Xizmat ko'rilmoqda"),
                        ('done',     'Yakunlandi'),
                        ('skipped',  "O'tkazib yuborildi"),
                    ],
                    default='waiting', max_length=20, verbose_name='Holat'
                )),
                ('room', models.CharField(default='MRT xonasi', max_length=50, verbose_name='Xona')),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('called_at',  models.DateTimeField(blank=True, null=True)),
                ('served_at',  models.DateTimeField(blank=True, null=True)),
                ('done_at',    models.DateTimeField(blank=True, null=True)),
                ('patient_card', models.ForeignKey(
                    on_delete=django.db.models.deletion.CASCADE,
                    related_name='queue_tickets',
                    to='patients.patientcard',
                    verbose_name='Bemor'
                )),
                ('service', models.ForeignKey(
                    blank=True, null=True,
                    on_delete=django.db.models.deletion.CASCADE,
                    related_name='queue_ticket',
                    to='services.patientservice',
                    verbose_name='Xizmat'
                )),
            ],
            options={
                'verbose_name': 'Navbat chipta',
                'verbose_name_plural': 'Navbat chiptalar',
                'ordering': ['ticket_number'],
            },
        ),
    ]