# Generated by Django 4.2.16 on 2024-12-09 09:04

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('invAPP', '0001_initial'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='inventory1c',
            name='date_of_decommission',
        ),
        migrations.AlterField(
            model_name='inventory1c',
            name='date_acceptance_accounting',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
    ]
