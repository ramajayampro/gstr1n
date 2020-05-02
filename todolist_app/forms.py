from django import forms
from todolist_app.models import TaskList
from todolist_app.models import Gstworker
from django.forms import ModelForm


class TaskForm(forms.ModelForm):
    class Meta:
       model = TaskList
       fields = ['task', 'done']
    

class StudentForm(forms.Form):  
    file      = forms.FileField(widget=forms.ClearableFileInput(attrs={'multiple': True})) # for creating file input

class GstForm(ModelForm):
    class Meta:
        model = Gstworker
        fields = ['GSTIN'] 

class FileFieldForm(forms.Form):
    file_field = forms.FileField(widget=forms.ClearableFileInput(attrs={'multiple': True}))      