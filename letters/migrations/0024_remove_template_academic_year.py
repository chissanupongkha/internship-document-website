from django.db import migrations


def merge_years(apps, schema_editor):
    """Without the academic year every letter type has one list of versions and one active version.

    Versions are renumbered 1, 2, 3 … (older year first), and when several were active
    only the most recently updated one stays active.
    """
    T = apps.get_model('letters', 'LetterTemplate')
    for lt in set(T.objects.values_list('letter_type', flat=True)):
        rows = list(T.objects.filter(letter_type=lt).order_by('academic_year', 'version', 'id'))
        for n, t in enumerate(rows, start=1):
            if t.version != n:
                t.version = n
                t.save(update_fields=['version'])
        active = [t for t in rows if t.is_active]
        if len(active) > 1:
            keep = max(active, key=lambda t: (t.updated_at, t.id))
            for t in active:
                if t.id != keep.id:
                    t.is_active = False
                    t.save(update_fields=['is_active'])


class Migration(migrations.Migration):

    dependencies = [('letters', '0023_course_instructor')]

    operations = [
        migrations.RunPython(merge_years, migrations.RunPython.noop),
        migrations.RemoveField(model_name='lettertemplate', name='academic_year'),
    ]
