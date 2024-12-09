from django.contrib import admin

# Register your models here.
from .models import Workers, Inventory1C, Characteristics, Relation3, Relation4, TypeOfEquipment

admin.site.register(TypeOfEquipment)
admin.site.register(Inventory1C)
admin.site.register(Workers)
admin.site.register(Characteristics)
admin.site.register(Relation3)
admin.site.register(Relation4)
