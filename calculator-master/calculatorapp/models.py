from django.db import models

# Create your models here.
class InData(models.Model):
    FST=models.CharField(max_length=250)

class InData2(models.Model):
    DESCRIPTION=models.CharField(max_length=250)
    NPV=models.DecimalField(max_digits=12, decimal_places=0)
    IRR=models.DecimalField(max_digits=5, decimal_places=2)
    PI=models.DecimalField(max_digits=5, decimal_places=2)
    SE=models.DecimalField(max_digits=12, decimal_places=0)
    BP=models.DecimalField(max_digits=12, decimal_places=0)
    RESERVES=models.DecimalField(max_digits=12, decimal_places=0)
    PROD_RATE=models.DecimalField(max_digits=12, decimal_places=0)
    ORE_TYPE=models.CharField(max_length=250)
    STORE_FILE=models.FileField(upload_to="XLSFILES")