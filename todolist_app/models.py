from django.db import models

# Create your models here.
class TaskList(models.Model):
    task = models.CharField(max_length=300)
    done = models.BooleanField(default=False)


    def __str__(self):
        return self.task + " - Task - " + str(self.done)

class Gstworker(models.Model):
    GSTIN = models.CharField(max_length=15)
    date = models.DateField(auto_now=True)
    r_counts = models.IntegerField(blank=True, null=True)   

    