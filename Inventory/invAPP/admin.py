from django.contrib import admin

# Register your models here.
from .models import TypeOfEquipment, Inventory1C, Workers

admin.site.register(TypeOfEquipment)
admin.site.register(Inventory1C)
admin.site.register(Workers)
