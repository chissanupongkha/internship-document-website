# -*- coding: utf-8 -*-
from django.db import models
from django.contrib.auth.models import User


class UploadedBatch(models.Model):
    """One uploaded Excel/CSV file."""
    file_name = models.CharField(max_length=255)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)

    def __str__(self):
        uploader = self.uploaded_by.username if self.uploaded_by else "System"
        return f"Batch #{self.id} - {self.file_name} (Uploaded by {uploader})"


class BatchUpload(models.Model):
    """Legacy or alternate batch tracking model."""
    filename = models.CharField(max_length=255)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    total_records = models.IntegerField(default=0)

    def __str__(self):
        return f"Batch #{self.id} - {self.filename}"


class StudentRecord(models.Model):
    batch = models.ForeignKey('UploadedBatch', on_delete=models.CASCADE, related_name='records')
    row_index = models.IntegerField(default=0)
    
    # Core Tracking Keys
    student_id = models.CharField(max_length=50)
    display_name = models.CharField(max_length=255)
    letter_type = models.CharField(max_length=100)
    course_id = models.CharField(max_length=255, blank=True, default='')   # e.g., "ITCS499"
    semester = models.CharField(max_length=20, blank=True, default='')    # e.g., "2026/1"
    times = models.PositiveIntegerField(default=1)                       # 1st time, 2nd time, etc.
    
    # Details
    year = models.CharField(max_length=50, blank=True, default='')
    company = models.CharField(max_length=255, blank=True)
    contractor = models.CharField(max_length=255, blank=True)
    contractor_position = models.CharField(max_length=255, blank=True)
    internship_position = models.CharField(max_length=255, blank=True)
    period = models.CharField(max_length=255, blank=True)
    data = models.JSONField(default=dict)
    ref_no = models.CharField(max_length=20, blank=True, default='')    # the running number in "อว 78.363 / 012"

    class Meta:
        ordering = ['-id']

    def __str__(self):
        return f"{self.student_id} - {self.display_name} ({self.course_id} {self.semester} #{self.times})"

    def get_ref_no(self):
        """The letter's reference number. A record that has none yet gets the next running number
        (001, 002, ...); staff can change it afterwards on the record page."""
        if self.ref_no:
            return self.ref_no
        from django.db import transaction
        with transaction.atomic():
            LetterSettings.objects.select_for_update().filter(pk=1).first()      # one allocation at a time
            used = [int(v) for v in StudentRecord.objects.exclude(ref_no='').values_list('ref_no', flat=True)
                    if v.isdigit()]
            self.ref_no = f"{(max(used) + 1 if used else 1):03d}"
            self.save(update_fields=['ref_no'])
        return self.ref_no


class Logo(models.Model):
    """A letter logo staff can choose for a document template."""
    name = models.CharField(max_length=100)
    image = models.ImageField(upload_to='logos/')
    uploaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


class LetterTemplate(models.Model):
    logo = models.ForeignKey('Logo', null=True, blank=True, on_delete=models.SET_NULL, related_name='templates')
    name = models.CharField(max_length=150)
    letter_type = models.CharField(max_length=50)                 # several versions of one type are allowed
    version = models.PositiveIntegerField(default=1)
    file = models.FileField(upload_to='letter_templates/')
    is_active = models.BooleanField(default=True)
    updated_at = models.DateTimeField(auto_now=True)

    @property
    def label(self):
        return f"{self.name} v{self.version}"

    def __str__(self):
        return f"{self.name} ({self.letter_type})"

class LetterType(models.Model):
    slug = models.SlugField(max_length=50, unique=True)
    display_name = models.CharField(max_length=150)

    def __str__(self):
        return self.display_name


class LetterTypeMapping(models.Model):
    raw_text = models.CharField(max_length=255, unique=True)
    letter_type = models.ForeignKey(LetterType, related_name='mappings', on_delete=models.CASCADE)

    def __str__(self):
        return f"{self.raw_text} → {self.letter_type.slug}"


class GeneratedDocument(models.Model):
    STATUS_GENERATED = 'generated'
    STATUS_SKIPPED = 'skipped'
    STATUS_SENT = 'sent'
    STATUS_CHOICES = [
        (STATUS_GENERATED, 'Generated'),
        (STATUS_SKIPPED, 'Skipped'),
        (STATUS_SENT, 'Sent'),
    ]

    record = models.ForeignKey(StudentRecord, related_name='documents', on_delete=models.CASCADE)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES)
    file_name = models.CharField(max_length=255, blank=True)
    file = models.FileField(upload_to='generated_letters/', blank=True, null=True)
    skip_reason = models.CharField(max_length=255, blank=True)
    ref_no = models.CharField(max_length=20, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    signed_file = models.FileField(upload_to='signed_docs/', null=True, blank=True)
    version = models.PositiveIntegerField(default=1)             # only the latest version is kept
    updated_at = models.DateTimeField(auto_now=True)
    email_sent_at = models.DateTimeField(null=True, blank=True)
    signed_at = models.DateTimeField(null=True, blank=True)      # when the signed PDF was uploaded

    def __str__(self):
        student_name = self.record.display_name if self.record else "Unknown Student"
        student_id = self.record.student_id if self.record else "N/A"
        letter_type = self.record.letter_type if self.record else "N/A"
        identifier = self.file_name or self.skip_reason or f"Ref: {self.ref_no}"
        status_label = self.get_status_display()
        return f"{student_name} ({student_id}) | {identifier} [{letter_type}] ({status_label})"


class GenerationBatch(models.Model):
    batch = models.ForeignKey(UploadedBatch, related_name='generation_runs', on_delete=models.CASCADE)
    created_at = models.DateTimeField(auto_now_add=True)
    documents = models.ManyToManyField(GeneratedDocument, related_name='generation_batches')


class AuditLog(models.Model):
    ACTION_CHOICES = [
        ('UPLOAD_EXCEL', 'Uploaded Excel Batch'),
        ('GENERATE_DOCS', 'Generated Draft Letters'),
        ('UPLOAD_SIGNED', 'Uploaded Signed PDFs'),
        ('SEND_EMAIL', 'Dispatched Emails'),
        ('EDIT_RECORD', 'Edited Student Record'),
        ('ADD_RECORD', 'Manually Added Student Record'),
    ]

    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    action = models.CharField(max_length=50, choices=ACTION_CHOICES)
    details = models.TextField(blank=True)
    ip_address = models.GenericIPAddressField(null=True, blank=True)
    timestamp = models.DateTimeField(auto_now_add=True)


class TemplateFieldMapping(models.Model):
    SOURCE_FIELD = 'field'
    SOURCE_STATIC = 'static'
    SOURCE_UNMAPPED = 'unmapped'
    SOURCE_CHOICES = [
        (SOURCE_FIELD, 'Data Field'),
        (SOURCE_STATIC, 'Static Text'),
        (SOURCE_UNMAPPED, 'Not Mapped Yet'),
    ]

    template = models.ForeignKey('LetterTemplate', related_name='field_mappings', on_delete=models.CASCADE)
    placeholder = models.CharField(max_length=100)
    source_type = models.CharField(max_length=10, choices=SOURCE_CHOICES, default=SOURCE_UNMAPPED)
    source_value = models.CharField(max_length=255, blank=True, default='')

    class Meta:
        unique_together = ('template', 'placeholder')


class StudentLetter(models.Model):
    student_id = models.CharField(max_length=20)
    display_name = models.CharField(max_length=255, blank=True, null=True)
    letter_type = models.CharField(max_length=50)
    company = models.CharField(max_length=255, blank=True, null=True)
    course_id = models.CharField(max_length=100, blank=True, null=True)
    data = models.JSONField(default=dict, blank=True)
    signed_file = models.FileField(upload_to='signed_letters/', blank=True, null=True)
    status = models.CharField(max_length=20, default='pending')
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.student_id} - {self.display_name or ''} ({self.letter_type})"


class EmailTemplate(models.Model):
    """Message sent with a signed letter. pk=1 with no letter type is the default for every type;
    a row with a letter type is used only for that type."""
    DEFAULT_SUBJECT = "Signed Internship Letter & Draft - {name}"
    DEFAULT_BODY = (
        "Dear {name},\n\n"
        "Please find attached your official signed internship letter (PDF) "
        "and draft copy.\n\n"
        "Best regards,\n"
        "{staff_name}\n"
        "Faculty of ICT, Mahidol University\n"
        "Email: {staff_email}"
    )

    letter_type = models.OneToOneField('LetterType', null=True, blank=True, on_delete=models.CASCADE,
                                       related_name='email_template')
    subject = models.CharField(max_length=255, default=DEFAULT_SUBJECT)
    body = models.TextField(default=DEFAULT_BODY)
    updated_at = models.DateTimeField(auto_now=True)
    updated_by = models.ForeignKey(User, null=True, blank=True, on_delete=models.SET_NULL)

    @classmethod
    def load(cls, letter_type=None):
        """The template for this letter type, or the default one when the type has none."""
        if letter_type is not None:
            own = cls.objects.filter(letter_type=letter_type).first()
            if own:
                return own
        obj = cls.objects.filter(letter_type__isnull=True).order_by('pk').first()
        if obj is None:
            obj = cls.objects.create()
        return obj


# ----------------------------------------------------------------------
# Student / subject roster (used by the "Request Status" pages)
# ----------------------------------------------------------------------
class Course(models.Model):
    """A subject, e.g. ITCS499. Shown as 'Subject' in the UI."""
    code = models.CharField(max_length=30, unique=True)          # stored upper-case, no spaces
    name = models.CharField(max_length=255, blank=True)
    name_th = models.CharField(max_length=255, blank=True)
    min_year = models.PositiveSmallIntegerField(default=2)       # lowest study year allowed to register (year 1 not allowed)
    semester = models.CharField(max_length=20, blank=True)       # e.g. "2569/1"
    start_date = models.DateField(null=True, blank=True)         # once reached, the subject can no longer be deleted
    end_date = models.DateField(null=True, blank=True)           # last day of the semester (the semester has expired after it)
    instructor_name = models.CharField("Supervising instructor", max_length=200, blank=True)
    instructor_email = models.CharField(max_length=200, blank=True)
    instructor_phone = models.CharField(max_length=100, blank=True)

    @property
    def label(self):
        """2026/1 ITCS371 Internship - shown everywhere instead of the bare code."""
        return " ".join(x for x in (self.semester, self.code, self.name) if x)

    @property
    def ended(self):
        from django.utils import timezone
        return bool(self.end_date and self.end_date < timezone.localdate())

    @property
    def phase(self):
        """'upcoming' / 'open' / 'ended' / '' (no dates)."""
        if self.ended:
            return 'ended'
        if self.started:
            return 'open'
        return 'upcoming' if self.start_date else ''

    @property
    def started(self):
        from django.utils import timezone
        return bool(self.start_date and self.start_date <= timezone.localdate())

    class Meta:
        ordering = ['code']

    def __str__(self):
        return f"{self.code} {self.name}".strip()


class Student(models.Model):
    student_id = models.CharField(max_length=50, unique=True)
    display_name = models.CharField(max_length=255)
    email = models.EmailField(blank=True)
    year = models.CharField(max_length=50, blank=True)           # study year, e.g. "3"

    class Meta:
        ordering = ['student_id']

    def __str__(self):
        return f"{self.student_id} {self.display_name}"


class Enrollment(models.Model):
    """A student registered in a subject in a semester."""
    student = models.ForeignKey(Student, related_name='enrollments', on_delete=models.CASCADE)
    course = models.ForeignKey(Course, related_name='enrollments', on_delete=models.CASCADE)
    semester = models.CharField(max_length=20)                   # e.g. "2568/1"

    class Meta:
        unique_together = ('student', 'course', 'semester')
        ordering = ['-semester', 'course__code']

    def __str__(self):
        return f"{self.student.student_id} - {self.course.code} ({self.semester})"



class LetterSettings(models.Model):
    """One shared record (pk=1): values printed on every letter, editable by staff."""
    ref_prefix = models.CharField("Reference number prefix", max_length=50, default='อว 78.363')
    letter_date = models.DateField("Fixed letter date", null=True, blank=True,
                                   help_text="Leave empty to use the date the letter is generated.")
    course_contact = models.TextField("Course supervisor line", blank=True, default=(
        "อาจารย์ผู้ดูแลรายวิชา อาจารย์ผศ. ดร.ศิริเพ็ญ พงษ์ไพเชฐ "
        "อีเมล์ siripen.pon@mahidol.ac.th หรือ โทรศัพท์ 083-0610496"))
    dean_name = models.CharField(max_length=150, default="(ดร.พัฒนศักดิ์ มงคลวัฒน์)")
    dean_title = models.CharField(max_length=200, default="คณบดีคณะเทคโนโลยีสารสนเทศและการสื่อสาร")
    line_spacing = models.DecimalField(max_digits=3, decimal_places=2, default=1.0)
    logo_width_cm = models.DecimalField(max_digits=3, decimal_places=1, default=2.0)
    header_gap_pt = models.PositiveSmallIntegerField(default=2)
    logo_align = models.CharField(max_length=6, default='center')
    header_align = models.CharField(max_length=6, default='right')
    header_font_pt = models.DecimalField(max_digits=3, decimal_places=1, default=14)
    logo_gap_pt = models.PositiveSmallIntegerField(default=0)
    updated_at = models.DateTimeField(auto_now=True)

    @classmethod
    def load(cls):
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj


class StaffProfile(models.Model):
    """Role and account expiry for a staff login. Added date is User.date_joined."""
    ROLE_ADMIN = 'admin'
    ROLE_STAFF = 'staff'
    ROLE_VIEWER = 'viewer'
    ROLE_CHOICES = [
        (ROLE_ADMIN, 'Admin: everything, including accounts and configuration'),
        (ROLE_STAFF, 'Staff: requests, tracking, academic data and reports'),
        (ROLE_VIEWER, 'Viewer: look only, no changes'),
    ]
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile')
    role = models.CharField(max_length=10, choices=ROLE_CHOICES, default=ROLE_STAFF)
    expires_on = models.DateField(null=True, blank=True)

    @property
    def role_label(self):
        return self.get_role_display().split(':')[0]

    @property
    def expired(self):
        from django.utils import timezone
        return bool(self.expires_on and self.expires_on < timezone.localdate())


class ChangeLog(models.Model):
    """History of admin edits to the company / internship period of a request."""
    record = models.ForeignKey(StudentRecord, related_name='changes', on_delete=models.CASCADE)
    field = models.CharField(max_length=30)                      # company | period
    old_value = models.CharField(max_length=255, blank=True)
    new_value = models.CharField(max_length=255, blank=True)
    reason = models.CharField(max_length=255, blank=True)        # e.g. "approved on paper, ref. no. ..."
    changed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    changed_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-changed_at', '-id']

    def __str__(self):
        return f"{self.record_id} {self.field}: {self.old_value} -> {self.new_value}"
