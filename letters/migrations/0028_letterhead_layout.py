from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [('letters', '0027_letterhead_settings')]
    operations = [
        migrations.AddField(model_name='lettersettings', name='logo_align', field=models.CharField(default='center', max_length=6)),
        migrations.AddField(model_name='lettersettings', name='header_align', field=models.CharField(default='right', max_length=6)),
        migrations.AddField(model_name='lettersettings', name='header_font_pt', field=models.DecimalField(decimal_places=1, default=14, max_digits=3)),
        migrations.AddField(model_name='lettersettings', name='logo_gap_pt', field=models.PositiveSmallIntegerField(default=0)),
    ]
