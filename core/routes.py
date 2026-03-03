from django.urls import path
from . import views

urlpatterns = [
    path('admin_dashboard/', views.admin_dashboard, name='admin_dashboard'),
    
    # Cheque lists by status
    path('cheques/', views.cheque_list, name='cheque_list'),
    path('cheques/deposited/', views.cheque_deposited_list, name='cheque_deposited_list'),
    path('cheques/rejected/', views.cheque_rejected_list, name='cheque_rejected_list'),
    
    # CRUD operations
    path('cheques/create/', views.cheque_create, name='cheque_create'),
    path('cheques/update/<int:pk>/', views.cheque_update, name='cheque_update'),
    path('cheques/delete/<int:pk>/', views.cheque_delete, name='cheque_delete'),
    
    # Status change operations
    path('cheques/reject/<int:pk>/', views.cheque_reject, name='cheque_reject'),
    path('cheques/deposit/<int:pk>/', views.cheque_deposit, name='cheque_deposit'),
    path('cheques/revert/<int:pk>/', views.cheque_revert_to_new, name='cheque_revert'),


    # Reports
    path('reports/', views.reports_page, name='reports_page'),
    path('reports/export/', views.export_report_excel, name='export_report_excel'),
    path('api/company-suggestions/', views.get_company_suggestions, name='company_suggestions'),

    # Email Deposited Cheques
    path('email-deposited/', views.email_deposited_cheques, name='email_deposited_cheques'),
    path('email-deposited/send/', views.send_deposited_email, name='send_deposited_email'),
    path('email-deposited/preview/', views.get_selected_cheques_data, name='get_selected_cheques_data'),
]