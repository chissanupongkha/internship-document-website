# -*- coding: utf-8 -*-
"""Fills the database with realistic, clearly fictional demo data for demos and screenshots.

    python manage.py seed_demo          # add demo data
    python manage.py seed_demo --reset  # remove earlier demo data first
"""
import datetime

from django.core.management.base import BaseCommand

from letters.models import (Course, Enrollment, GeneratedDocument, Student,
                            StudentRecord, UploadedBatch)

COMPANIES = [
    ('บริษัท สยามดิจิทัลโซลูชัน จำกัด', 'คุณสุดา ใจดี', 'ผู้จัดการฝ่ายบุคคล'),
    ('บริษัท เจ้าพระยาซอฟต์แวร์ จำกัด', 'คุณวิชัย แสงทอง', 'หัวหน้าฝ่ายพัฒนาระบบ'),
    ('บริษัท กรุงเทพดาต้าแล็บ จำกัด', 'คุณมาลี รุ่งเรือง', 'ผู้จัดการฝ่ายทรัพยากรบุคคล'),
    ('บริษัท นนทบุรีคลาวด์ จำกัด', 'คุณธีรพงษ์ สุขสวัสดิ์', 'ผู้อำนวยการฝ่ายเทคโนโลยี'),
    ('บริษัท ปทุมธานีอินโนเวชัน จำกัด', 'คุณพิมพ์ใจ ศรีสุข', 'ผู้จัดการฝ่ายบุคคล'),
]
PEOPLE = [
    ('6588001', 'นาย', 'สมชาย', 'รักไทย', 3), ('6588014', 'นางสาว', 'กมลชนก', 'พึ่งบุญ', 3),
    ('6588027', 'นาย', 'ธนกฤต', 'วงศ์สว่าง', 3), ('6588033', 'นางสาว', 'ปิยะนุช', 'เจริญสุข', 3),
    ('6588045', 'นาย', 'อภิชาติ', 'มั่นคง', 3), ('6588052', 'นางสาว', 'ชลธิชา', 'ทองดี', 4),
    ('6588068', 'นาย', 'พงศกร', 'แก้วมณี', 4), ('6588071', 'นางสาว', 'วรรณา', 'สุขใจ', 4),
]
POSITIONS = ['นักพัฒนาซอฟต์แวร์', 'นักวิเคราะห์ข้อมูล', 'วิศวกรระบบคลาวด์', 'นักออกแบบ UX/UI', 'ผู้ทดสอบซอฟต์แวร์']
PERIODS = ['2 มิถุนายน - 25 กรกฎาคม 2569', '9 มิถุนายน - 1 สิงหาคม 2569']
TYPES = ['Request Letter for Internship Placement', 'Student Referral Letter']


class Command(BaseCommand):
    help = "Add realistic fictional demo data (students, subject, semester, requests)."

    def add_arguments(self, parser):
        parser.add_argument('--reset', action='store_true')

    def handle(self, *args, **opts):
        if opts['reset']:
            UploadedBatch.objects.filter(file_name='demo_requests.csv').delete()
            Student.objects.filter(student_id__in=[p[0] for p in PEOPLE]).delete()
        course, _ = Course.objects.get_or_create(code='ITCS499', defaults={
            'name': 'Internship', 'name_th': 'การฝึกงาน', 'semester': '2569/1',
            'start_date': datetime.date(2026, 6, 1)})
        batch = UploadedBatch.objects.create(file_name='demo_requests.csv')
        for i, (sid, pre, first, last, year) in enumerate(PEOPLE):
            stu, _ = Student.objects.get_or_create(student_id=sid, defaults={
                'display_name': f'{pre}{first} {last}', 'year': str(year), 'email': f'{sid}@student.example.edu'})
            Enrollment.objects.get_or_create(student=stu, course=course, semester='2569/1')
            if i == 7:
                continue                                  # one student has not asked yet
            company, contact, position = COMPANIES[i % len(COMPANIES)]
            period = PERIODS[i % 2]
            rec = StudentRecord.objects.create(
                batch=batch, row_index=i, student_id=sid, display_name=f'{pre} {first} {last}',
                letter_type=TYPES[0 if i % 3 else 1], course_id='ITCS499 การฝึกงาน', semester='2569/1',
                year=str(year), company=company, contractor=contact, contractor_position=position,
                internship_position=POSITIONS[i % len(POSITIONS)], period=period,
                data={'รหัสนักศึกษา': sid, 'คำนำหน้าชื่อ': pre, 'ชื่อ (ภาษาไทย)': first, 'นามสกุล': last,
                      'ชั้นปีที่': str(year), 'ชื่อหน่วยงาน/บริษัท ที่ต้องการฝึกงาน': company,
                      'โดยทำหนังสือถึงคุณ (สอบถามผู้ติดต่อฝั่งบริษัท)': contact,
                      'ตำแหน่งของการฝึกงาน': POSITIONS[i % len(POSITIONS)], 'ช่วงเวลาของการฝึกงาน': period,
                      'โปรดเลือกรายวิชาของการไปฝึกงาน': 'ITCS499 การฝึกงาน', 'ประเภทหนังสือที่ขอ': TYPES[0 if i % 3 else 1]})
            if i < 5:
                GeneratedDocument.objects.create(record=rec, status=GeneratedDocument.STATUS_GENERATED,
                                                 ref_no=str(2600 + i), file_name=f'{sid}_demo.docx')
        self.stdout.write(self.style.SUCCESS(f"Demo data added: {len(PEOPLE)} students, batch #{batch.id}."))
