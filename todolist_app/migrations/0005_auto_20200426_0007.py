# Generated by Django 3.0.4 on 2020-04-25 18:37

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('todolist_app', '0004_gstworker_r_count'),
    ]

    operations = [
        migrations.RenameField(
            model_name='gstworker',
            old_name='r_count',
            new_name='r_counts',
        ),
    ]
