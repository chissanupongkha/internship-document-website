from decimal import Decimal

from django.db import migrations, models


def tighten(apps, schema_editor):
    """Built-in letters should fit on one page: move settings still on the old default (1.15) to 1.0.
    A value the staff chose themselves is left alone."""
    LetterSettings = apps.get_model('letters', 'LetterSettings')
    LetterSettings.objects.filter(line_spacing=Decimal('1.15')).update(line_spacing=Decimal('1.00'))


class Migration(migrations.Migration):

    dependencies = [
        ('letters', '0024_remove_template_academic_year'),
    ]

    operations = [
        migrations.AlterField(
            model_name='lettersettings',
            name='line_spacing',
            field=models.DecimalField(decimal_places=2, default=1.0, max_digits=3),
        ),
        migrations.RunPython(tighten, migrations.RunPython.noop),
    ]
