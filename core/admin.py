from django.contrib import admin
from .models import ChequeModel


@admin.register(ChequeModel)
class ChequeModelAdmin(admin.ModelAdmin):
    list_display = ['company_name', 'cheque_no', 'amount', 'cheque_date', 'created_at']
    search_fields = ['company_name', 'cheque_no']