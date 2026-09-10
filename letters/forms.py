from django import forms
from docxtpl import DocxTemplate
from .models import LetterTemplate, StudentRecord


class LetterTemplateForm(forms.ModelForm):
    class Meta:
        model = LetterTemplate
        fields = ['name', 'letter_type', 'file']
        widgets = {
            'name': forms.TextInput(attrs={
                'class': 'w-full border border-gray-300 rounded-lg p-2.5 text-sm focus:ring-[#006C6C] focus:border-[#006C6C]',
                'placeholder': 'e.g., Mahidol ICT Request Letter 2026'
            }),
            'letter_type': forms.Select(attrs={
                'class': 'w-full border border-gray-300 rounded-lg p-2.5 text-sm focus:ring-[#006C6C] focus:border-[#006C6C]'
            }),
            'file': forms.FileInput(attrs={
                'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-[#006C6C]/10 file:text-[#006C6C] hover:file:bg-[#006C6C]/20',
                'accept': '.docx'
            }),
        }

    def clean_file(self):
        file = self.cleaned_data.get('file')
        if file:
            if not file.name.lower().endswith('.docx'):
                raise forms.ValidationError("Only Microsoft Word (.docx) files are supported.")
            try:
                doc = DocxTemplate(file)
                doc.get_undeclared_variables()
            except Exception as e:
                raise forms.ValidationError(f"Invalid DOCX structure or corrupt template: {str(e)}")
        return file


class StudentRecordForm(forms.ModelForm):
    class Meta:
        model = StudentRecord
        fields = [
            'student_id', 'display_name', 'letter_type', 'course_id',
            'semester', 'company', 'contractor', 'contractor_position',
            'internship_position', 'period'
        ]
        widgets = {
            'student_id': forms.TextInput(attrs={'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm'}),
            'display_name': forms.TextInput(attrs={'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm'}),
            'letter_type': forms.TextInput(attrs={'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm'}),
            'course_id': forms.TextInput(attrs={'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm'}),
            'semester': forms.TextInput(attrs={'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm'}),
            'company': forms.TextInput(attrs={'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm'}),
            'contractor': forms.TextInput(attrs={'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm'}),
            'contractor_position': forms.TextInput(attrs={'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm'}),
            'internship_position': forms.TextInput(attrs={'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm'}),
            'period': forms.TextInput(attrs={'class': 'w-full border border-gray-300 rounded-lg p-2 text-sm'}),
        }