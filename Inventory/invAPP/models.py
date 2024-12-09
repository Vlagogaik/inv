from django.db import models

class Characteristics(models.Model):
    id_char = models.AutoField(primary_key=True)
    char_name = models.TextField()

    def __str__(self):
        return self.char_name

class TypeOfEquipment(models.Model):
    id_type = models.AutoField(primary_key=True)
    type_name = models.TextField()

    def __str__(self):
        return self.type_name

class Workers(models.Model):
    id_workers = models.AutoField(primary_key=True)
    full_name = models.TextField()
    department = models.TextField()

    def __str__(self):
        return self.full_name

class Inventory1C(models.Model):
    id = models.AutoField(primary_key=True)
    accounting_name = models.TextField()
    real_name = models.TextField()
    inventory_decimal = models.TextField()
    # date_acceptance_accounting = models.DateField()
    date_acceptance_accounting = models.CharField(max_length=100, null=True, blank=True)
    date_of_decommission = models.CharField(max_length=100, null=False, default="")
    # date_of_decommission = models.DateField(null=True, blank=True)
    initial_cost = models.DecimalField(max_digits=100, decimal_places=2)
    id_workers = models.ForeignKey(Workers, on_delete=models.CASCADE)
    id_type = models.ForeignKey(TypeOfEquipment, on_delete=models.CASCADE)

    def __str__(self):
        return self.accounting_name

class Relation3(models.Model):
    type_of_equipment_id_type = models.ForeignKey(TypeOfEquipment, on_delete=models.CASCADE)
    characteristics_id_char = models.ForeignKey(Characteristics, on_delete=models.CASCADE)

    class Meta:
        unique_together = ('type_of_equipment_id_type', 'characteristics_id_char')

    def __str__(self):
        return f"Rel {self.type_of_equipment_id_type} - {self.characteristics_id_char}"

class Relation4(models.Model):
    inventory_1c_id = models.ForeignKey(Inventory1C, on_delete=models.CASCADE)
    characteristics_id_char = models.ForeignKey(Characteristics, on_delete=models.CASCADE)
    value = models.TextField()

    class Meta:
        unique_together = ('inventory_1c_id', 'characteristics_id_char')

    def __str__(self):
        return f"Rel {self.inventory_1c_id} - {self.characteristics_id_char}"
