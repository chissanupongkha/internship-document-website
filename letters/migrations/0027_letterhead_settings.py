from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [('letters', '0026_record_ref_no')]
    operations = [
        migrations.AddField(model_name='lettersettings', name='logo_width_cm',
                            field=models.DecimalField(decimal_places=1, default=2.0, max_digits=3)),
        migrations.AddField(model_name='lettersettings', name='header_gap_pt',
                            field=models.PositiveSmallIntegerField(default=2)),
    ]
