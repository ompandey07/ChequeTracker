from .models import ChequeModel, DepositModel ,  KhajaBill, KhajaBillItem
from django.db.models.functions import TruncMonth, TruncDate
from django.contrib.auth.decorators import login_required
from django.views.decorators.http import require_POST
from django.http import JsonResponse, HttpResponse
from datetime import datetime, timedelta , date
from django.shortcuts import render, redirect
from django.core.paginator import Paginator
from django.db.models import Sum, Count, Q
from django.core.mail import EmailMessage
from django.db.models import Sum, Count
from django.http import JsonResponse
from django.conf import settings
from decimal import Decimal
from io import BytesIO
import pandas as pd
import json




# !=========================================== ADMIN DASHBOARD ===========================================!
@login_required(login_url='/auth/login/')
def admin_dashboard(request):
    # Basic counts
    total_cheques = ChequeModel.objects.count()
    new_cheques = ChequeModel.objects.filter(status='New').count()
    deposited_cheques = ChequeModel.objects.filter(status='Deposited').count()
    rejected_cheques = ChequeModel.objects.filter(status='Rejected').count()
    
    # Amount calculations
    total_amount = ChequeModel.objects.aggregate(total=Sum('amount'))['total'] or 0
    new_amount = ChequeModel.objects.filter(status='New').aggregate(total=Sum('amount'))['total'] or 0
    deposited_amount = ChequeModel.objects.filter(status='Deposited').aggregate(total=Sum('amount'))['total'] or 0
    rejected_amount = ChequeModel.objects.filter(status='Rejected').aggregate(total=Sum('amount'))['total'] or 0
    
    # Recent cheques (last 5)
    recent_cheques = ChequeModel.objects.all()[:5]
    
    # Monthly data for chart (last 6 months)
    six_months_ago = datetime.now() - timedelta(days=180)
    monthly_data = ChequeModel.objects.filter(
        created_at__gte=six_months_ago
    ).annotate(
        month=TruncMonth('created_at')
    ).values('month').annotate(
        count=Count('id'),
        total=Sum('amount')
    ).order_by('month')
    
    # Prepare chart data
    months = []
    monthly_counts = []
    monthly_amounts = []
    for item in monthly_data:
        months.append(item['month'].strftime('%b %Y'))
        monthly_counts.append(item['count'])
        monthly_amounts.append(float(item['total']) if item['total'] else 0)
    
    # Daily data for last 7 days
    seven_days_ago = datetime.now() - timedelta(days=7)
    daily_data = ChequeModel.objects.filter(
        created_at__gte=seven_days_ago
    ).annotate(
        day=TruncDate('created_at')
    ).values('day').annotate(
        count=Count('id')
    ).order_by('day')
    
    days = []
    daily_counts = []
    for item in daily_data:
        days.append(item['day'].strftime('%d %b'))
        daily_counts.append(item['count'])
    
    # Top companies by amount
    top_companies = ChequeModel.objects.values('company_name').annotate(
        total=Sum('amount'),
        count=Count('id')
    ).order_by('-total')[:5]
    
    context = {
        'total_cheques': total_cheques,
        'new_cheques': new_cheques,
        'deposited_cheques': deposited_cheques,
        'rejected_cheques': rejected_cheques,
        'total_amount': total_amount,
        'new_amount': new_amount,
        'deposited_amount': deposited_amount,
        'rejected_amount': rejected_amount,
        'recent_cheques': recent_cheques,
        'top_companies': top_companies,
        # Chart data as JSON
        'months_json': json.dumps(months),
        'monthly_counts_json': json.dumps(monthly_counts),
        'monthly_amounts_json': json.dumps(monthly_amounts),
        'days_json': json.dumps(days),
        'daily_counts_json': json.dumps(daily_counts),
    }
    return render(request, 'Base/admin_dashboard.html', context)
# !=========================================== END OF ADMIN DASHBOARD ===========================================!




# !=========================================== CHEQUE VIEWS ===========================================!
@login_required(login_url='/auth/login/')
def cheque_list(request):
    """Show only NEW status cheques"""
    cheques = ChequeModel.objects.filter(status='New')
    paginator = Paginator(cheques, 15)
    page_number = request.GET.get('page', 1)
    page_obj = paginator.get_page(page_number)

    current = page_obj.number
    total = paginator.num_pages
    start_p = max(current - 2, 1)
    end_p = min(current + 2, total)
    page_range = range(start_p, end_p + 1)

    start_index = (page_obj.number - 1) * paginator.per_page

    context = {
        'page_obj': page_obj,
        'page_range': page_range,
        'start_index': start_index,
        'total_count': paginator.count,
    }
    return render(request, 'Core/cheques.html', context)


# ! Show only DEPOSITED status cheques
@login_required(login_url='/auth/login/')
def cheque_deposited_list(request):
    """Show only DEPOSITED status cheques"""
    cheques = ChequeModel.objects.filter(status='Deposited').select_related('deposit')
    paginator = Paginator(cheques, 15)
    page_number = request.GET.get('page', 1)
    page_obj = paginator.get_page(page_number)

    current = page_obj.number
    total = paginator.num_pages
    start_p = max(current - 2, 1)
    end_p = min(current + 2, total)
    page_range = range(start_p, end_p + 1)

    start_index = (page_obj.number - 1) * paginator.per_page

    context = {
        'page_obj': page_obj,
        'page_range': page_range,
        'start_index': start_index,
        'total_count': paginator.count,
    }
    return render(request, 'Core/cheques_deposited.html', context)


# ! Show only REJECTED status cheques 
@login_required(login_url='/auth/login/')
def cheque_rejected_list(request):
    """Show only REJECTED status cheques"""
    cheques = ChequeModel.objects.filter(status='Rejected')
    paginator = Paginator(cheques, 15)
    page_number = request.GET.get('page', 1)
    page_obj = paginator.get_page(page_number)

    current = page_obj.number
    total = paginator.num_pages
    start_p = max(current - 2, 1)
    end_p = min(current + 2, total)
    page_range = range(start_p, end_p + 1)

    start_index = (page_obj.number - 1) * paginator.per_page

    context = {
        'page_obj': page_obj,
        'page_range': page_range,
        'start_index': start_index,
        'total_count': paginator.count,
    }
    return render(request, 'Core/cheques_rejected.html', context)


# ! This view allows creating a new cheque, and it simply saves the provided cheque details to the database with a default status of "New".
@login_required(login_url='/auth/login/')
def cheque_create(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            ChequeModel.objects.create(
                company_name=data['company_name'],
                cheque_no=data['cheque_no'],
                amount=data['amount'],
                cheque_date=data['cheque_date'],
                remarks=data.get('remarks', ''),
                status='New',  # Default status
            )
            return JsonResponse({'success': True, 'message': 'Cheque saved successfully.'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    return JsonResponse({'success': False, 'message': 'Invalid request.'})


# ! This view allows updating a cheque, and it simply updates the specified cheque in the database.
@login_required(login_url='/auth/login/')
def cheque_update(request, pk):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            cheque = ChequeModel.objects.get(pk=pk)
            cheque.company_name = data['company_name']
            cheque.cheque_no = data['cheque_no']
            cheque.amount = data['amount']
            cheque.cheque_date = data['cheque_date']
            cheque.remarks = data.get('remarks', '')
            cheque.save()
            return JsonResponse({'success': True, 'message': 'Cheque updated successfully.'})
        except ChequeModel.DoesNotExist:
            return JsonResponse({'success': False, 'message': 'Cheque not found.'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    return JsonResponse({'success': False, 'message': 'Invalid request.'})


# ! This view allows deleting a cheque, and it simply removes the specified cheque from the database.
@login_required(login_url='/auth/login/')
def cheque_delete(request, pk):
    if request.method == 'POST':
        try:
            cheque = ChequeModel.objects.get(pk=pk)
            cheque.delete()
            return JsonResponse({'success': True, 'message': 'Cheque deleted successfully.'})
        except ChequeModel.DoesNotExist:
            return JsonResponse({'success': False, 'message': 'Cheque not found.'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    return JsonResponse({'success': False, 'message': 'Invalid request.'})


# ! This view allows marking a cheque as Rejected, and it simply updates the status of the specified cheque to "Rejected".
@login_required(login_url='/auth/login/')
def cheque_reject(request, pk):
    """Change status to Rejected"""
    if request.method == 'POST':
        try:
            cheque = ChequeModel.objects.get(pk=pk)
            cheque.status = 'Rejected'
            cheque.save()
            return JsonResponse({'success': True, 'message': 'Cheque marked as Rejected.'})
        except ChequeModel.DoesNotExist:
            return JsonResponse({'success': False, 'message': 'Cheque not found.'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    return JsonResponse({'success': False, 'message': 'Invalid request.'})


# ! This view allows marking a cheque as Deposited, and it also creates a corresponding deposit record with the provided details.
@login_required(login_url='/auth/login/')
def cheque_deposit(request, pk):
    """Change status to Deposited with deposit details"""
    if request.method == 'POST':
        try:
            cheque = ChequeModel.objects.get(pk=pk)
            
            bank_name = request.POST.get('bank_name')
            branch_name = request.POST.get('branch_name')
            deposit_slip = request.FILES.get('deposit_slip')
            
            if not all([bank_name, branch_name, deposit_slip]):
                return JsonResponse({
                    'success': False, 
                    'message': 'All fields are required.'
                })
            
            # Create deposit record
            DepositModel.objects.create(
                cheque=cheque,
                bank_name=bank_name,
                branch_name=branch_name,
                deposit_slip=deposit_slip,
            )
            
            # Update cheque status
            cheque.status = 'Deposited'
            cheque.save()
            
            return JsonResponse({
                'success': True, 
                'message': 'Cheque marked as Deposited successfully.'
            })
        except ChequeModel.DoesNotExist:
            return JsonResponse({'success': False, 'message': 'Cheque not found.'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    return JsonResponse({'success': False, 'message': 'Invalid request.'})


#! This view allows reverting a cheque back to New status, and if it was deposited, it also deletes the associated deposit record.
@login_required(login_url='/auth/login/')
def cheque_revert_to_new(request, pk):
    """Revert cheque status back to New"""
    if request.method == 'POST':
        try:
            cheque = ChequeModel.objects.get(pk=pk)
            
            # If it was deposited, delete the deposit record
            if hasattr(cheque, 'deposit'):
                cheque.deposit.delete()
            
            cheque.status = 'New'
            cheque.save()
            
            return JsonResponse({
                'success': True, 
                'message': 'Cheque reverted to New status.'
            })
        except ChequeModel.DoesNotExist:
            return JsonResponse({'success': False, 'message': 'Cheque not found.'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    return JsonResponse({'success': False, 'message': 'Invalid request.'})
#! ==============================================================================================================
# !=========================================== END OF CHEQUE VIEWS ===========================================!


@login_required(login_url='/auth/login/')
def reports_page(request):
    """Main reports page with filters - Only Deposited and Rejected cheques"""
    
    # Get filter parameters
    date_from = request.GET.get('date_from', '')
    date_to = request.GET.get('date_to', '')
    company_name = request.GET.get('company_name', '')
    status = request.GET.get('status', '')
    cheque_no = request.GET.get('cheque_no', '')
    
    # Check if any filter is applied
    has_filters = any([date_from, date_to, company_name, status, cheque_no])
    
    # Base queryset - ONLY Deposited and Rejected (exclude New)
    cheques = ChequeModel.objects.filter(
        status__in=['Deposited', 'Rejected']
    ).select_related('deposit')
    
    # Apply filters
    if date_from:
        cheques = cheques.filter(
            Q(deposit__deposited_at__date__gte=date_from) | 
            Q(status='Rejected', updated_at__date__gte=date_from)
        )
    if date_to:
        cheques = cheques.filter(
            Q(deposit__deposited_at__date__lte=date_to) | 
            Q(status='Rejected', updated_at__date__lte=date_to)
        )
    if company_name:
        cheques = cheques.filter(company_name__icontains=company_name)
    if status:
        cheques = cheques.filter(status=status)
    if cheque_no:
        cheques = cheques.filter(cheque_no__icontains=cheque_no)
    
    # Order by deposited date (most recent first)
    cheques = cheques.order_by('-deposit__deposited_at', '-updated_at')
    
    # Calculate statistics for filtered data
    total_count = cheques.count()
    total_amount = cheques.aggregate(total=Sum('amount'))['total'] or 0
    
    # Status breakdown
    deposited_count = cheques.filter(status='Deposited').count()
    deposited_amount = cheques.filter(status='Deposited').aggregate(total=Sum('amount'))['total'] or 0
    rejected_count = cheques.filter(status='Rejected').count()
    rejected_amount = cheques.filter(status='Rejected').aggregate(total=Sum('amount'))['total'] or 0
    
    # Get unique company names for dropdown
    all_companies = list(
        ChequeModel.objects.filter(status__in=['Deposited', 'Rejected'])
        .values_list('company_name', flat=True)
        .distinct()
        .order_by('company_name')
    )
    
    # Only show data if filters are applied
    if has_filters:
        # Pagination
        paginator = Paginator(cheques, 20)
        page_number = request.GET.get('page', 1)
        page_obj = paginator.get_page(page_number)
        
        current = page_obj.number
        total_pages = paginator.num_pages
        start_p = max(current - 2, 1)
        end_p = min(current + 2, total_pages)
        page_range = range(start_p, end_p + 1)
        
        start_index = (page_obj.number - 1) * paginator.per_page
    else:
        # No data on initial load
        page_obj = None
        page_range = []
        start_index = 0
        total_count = 0
        total_amount = 0
    
    # Current date for default date_to
    current_date = date.today().strftime('%Y-%m-%d')
    
    context = {
        'page_obj': page_obj,
        'page_range': page_range,
        'start_index': start_index,
        'total_count': total_count,
        'total_amount': total_amount,
        'deposited_count': deposited_count,
        'deposited_amount': deposited_amount,
        'rejected_count': rejected_count,
        'rejected_amount': rejected_amount,
        'all_companies_json': json.dumps(all_companies),
        'has_filters': has_filters,
        'current_date': current_date,
        # Filter values for form persistence
        'filter_date_from': date_from,
        'filter_date_to': date_to if date_to else current_date,
        'filter_company': company_name,
        'filter_status': status,
        'filter_cheque_no': cheque_no,
    }
    
    return render(request, 'Report/reports.html', context)


@login_required(login_url='/auth/login/')
def export_report_excel(request):
    """Export filtered data to Excel report"""
    
    # Get filter parameters
    date_from = request.GET.get('date_from', '')
    date_to = request.GET.get('date_to', '')
    company_name = request.GET.get('company_name', '')
    status = request.GET.get('status', '')
    cheque_no = request.GET.get('cheque_no', '')
    
    # Base queryset - ONLY Deposited and Rejected
    cheques = ChequeModel.objects.filter(
        status__in=['Deposited', 'Rejected']
    ).select_related('deposit')
    
    # Apply filters
    if date_from:
        cheques = cheques.filter(
            Q(deposit__deposited_at__date__gte=date_from) | 
            Q(status='Rejected', updated_at__date__gte=date_from)
        )
    if date_to:
        cheques = cheques.filter(
            Q(deposit__deposited_at__date__lte=date_to) | 
            Q(status='Rejected', updated_at__date__lte=date_to)
        )
    if company_name:
        cheques = cheques.filter(company_name__icontains=company_name)
    if status:
        cheques = cheques.filter(status=status)
    if cheque_no:
        cheques = cheques.filter(cheque_no__icontains=cheque_no)
    
    cheques = cheques.order_by('-deposit__deposited_at', '-updated_at')
    
    # Prepare data for Excel
    data = []
    for idx, cheque in enumerate(cheques, 1):
        # Get deposited date
        deposited_date = ''
        bank_name = ''
        branch = ''
        
        if cheque.status == 'Deposited' and hasattr(cheque, 'deposit') and cheque.deposit:
            deposited_date = cheque.deposit.deposited_at.strftime('%Y-%m-%d %H:%M') if cheque.deposit.deposited_at else ''
            bank_name = cheque.deposit.bank_name
            branch = cheque.deposit.branch_name
        elif cheque.status == 'Rejected':
            deposited_date = cheque.updated_at.strftime('%Y-%m-%d %H:%M') if cheque.updated_at else ''
        
        row = {
            'SN': idx,
            'Company Name': cheque.company_name,
            'Cheque No': cheque.cheque_no,
            'Amount (NPR)': float(cheque.amount),
            'Cheque Date': cheque.cheque_date.strftime('%Y-%m-%d') if cheque.cheque_date else '',
            'Deposited/Rejected Date': deposited_date,
            'Status': cheque.status,
            'Bank Name': bank_name,
            'Branch': branch,
            'Remarks': cheque.remarks or '',
        }
        data.append(row)
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Calculate summary statistics
    total_amount = cheques.aggregate(total=Sum('amount'))['total'] or 0
    deposited_amount = cheques.filter(status='Deposited').aggregate(total=Sum('amount'))['total'] or 0
    rejected_amount = cheques.filter(status='Rejected').aggregate(total=Sum('amount'))['total'] or 0
    
    # Create Excel file with formatting
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # ═══════════════ FORMATS ═══════════════
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': 'white',
            'bg_color': '#1e1e2f',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'font_color': '#1e1e2f',
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 11,
            'font_color': '#6b7280',
            'align': 'center',
            'valign': 'vcenter'
        })
        
        cell_format = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        number_format = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': 'NPR #,##0.00'
        })
        
        date_format = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        status_deposited_format = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#dcfce7',
            'font_color': '#16a34a',
            'bold': True
        })
        
        status_rejected_format = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#fee2e2',
            'font_color': '#dc2626',
            'bold': True
        })
        
        summary_label_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': '#374151',
            'bg_color': '#f3f4f6',
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        summary_value_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': '#1e1e2f',
            'bg_color': '#f3f4f6',
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': 'NPR #,##0.00'
        })
        
        total_label_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': 'white',
            'bg_color': '#1e1e2f',
            'border': 1,
            'align': 'right',
            'valign': 'vcenter'
        })
        
        total_value_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': 'white',
            'bg_color': '#1e1e2f',
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': 'NPR #,##0.00'
        })
        
        # ═══════════════ MAIN REPORT SHEET ═══════════════
        worksheet = workbook.add_worksheet('Cheque Report')
        
        # Set column widths
        worksheet.set_column('A:A', 6)   # SN
        worksheet.set_column('B:B', 25)  # Company Name
        worksheet.set_column('C:C', 15)  # Cheque No
        worksheet.set_column('D:D', 18)  # Amount
        worksheet.set_column('E:E', 14)  # Cheque Date
        worksheet.set_column('F:F', 18)  # Deposited Date
        worksheet.set_column('G:G', 12)  # Status
        worksheet.set_column('H:H', 20)  # Bank Name
        worksheet.set_column('I:I', 15)  # Branch
        worksheet.set_column('J:J', 25)  # Remarks
        
        # Title
        worksheet.merge_range('A1:J1', 'CHEQUE REPORT - DEPOSITED & REJECTED', title_format)
        worksheet.set_row(0, 30)
        
        # Subtitle with date range
        filter_text = f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        if date_from or date_to:
            filter_text += f' | Date Range: {date_from or "Start"} to {date_to or "End"}'
        if company_name:
            filter_text += f' | Company: {company_name}'
        if status:
            filter_text += f' | Status: {status}'
        
        worksheet.merge_range('A2:J2', filter_text, subtitle_format)
        worksheet.set_row(1, 20)
        
        # Empty row
        worksheet.set_row(2, 10)
        
        # Summary section
        worksheet.merge_range('A4:B4', 'SUMMARY', header_format)
        worksheet.merge_range('C4:D4', '', header_format)
        
        worksheet.write('A5', 'Total Records:', summary_label_format)
        worksheet.write('B5', len(data), workbook.add_format({
            'bold': True, 'font_size': 11, 'bg_color': '#f3f4f6', 'border': 1, 'align': 'center'
        }))
        worksheet.write('C5', 'Total Amount:', summary_label_format)
        worksheet.write('D5', float(total_amount), summary_value_format)
        
        worksheet.write('A6', 'Deposited:', summary_label_format)
        worksheet.write('B6', cheques.filter(status='Deposited').count(), workbook.add_format({
            'bold': True, 'font_size': 11, 'bg_color': '#dcfce7', 'border': 1, 'align': 'center', 'font_color': '#16a34a'
        }))
        worksheet.write('C6', 'Deposited Amount:', summary_label_format)
        worksheet.write('D6', float(deposited_amount), summary_value_format)
        
        worksheet.write('A7', 'Rejected:', summary_label_format)
        worksheet.write('B7', cheques.filter(status='Rejected').count(), workbook.add_format({
            'bold': True, 'font_size': 11, 'bg_color': '#fee2e2', 'border': 1, 'align': 'center', 'font_color': '#dc2626'
        }))
        worksheet.write('C7', 'Rejected Amount:', summary_label_format)
        worksheet.write('D7', float(rejected_amount), summary_value_format)
        
        # Empty row
        worksheet.set_row(7, 15)
        
        # Data header row
        headers = ['SN', 'Company Name', 'Cheque No', 'Amount (NPR)', 'Cheque Date', 
                   'Deposited/Rejected Date', 'Status', 'Bank Name', 'Branch', 'Remarks']
        
        for col, header in enumerate(headers):
            worksheet.write(8, col, header, header_format)
        worksheet.set_row(8, 25)
        
        # Data rows
        row_num = 9
        for item in data:
            worksheet.write(row_num, 0, item['SN'], date_format)
            worksheet.write(row_num, 1, item['Company Name'], cell_format)
            worksheet.write(row_num, 2, item['Cheque No'], cell_format)
            worksheet.write(row_num, 3, item['Amount (NPR)'], number_format)
            worksheet.write(row_num, 4, item['Cheque Date'], date_format)
            worksheet.write(row_num, 5, item['Deposited/Rejected Date'], date_format)
            
            # Status with conditional formatting
            status_val = item['Status']
            if status_val == 'Deposited':
                worksheet.write(row_num, 6, status_val, status_deposited_format)
            else:
                worksheet.write(row_num, 6, status_val, status_rejected_format)
            
            worksheet.write(row_num, 7, item['Bank Name'], cell_format)
            worksheet.write(row_num, 8, item['Branch'], cell_format)
            worksheet.write(row_num, 9, item['Remarks'], cell_format)
            
            worksheet.set_row(row_num, 18)
            row_num += 1
        
        # Total row
        if data:
            worksheet.write(row_num, 0, '', total_label_format)
            worksheet.write(row_num, 1, '', total_label_format)
            worksheet.write(row_num, 2, 'TOTAL:', total_label_format)
            worksheet.write(row_num, 3, float(total_amount), total_value_format)
            for col in range(4, 10):
                worksheet.write(row_num, col, '', total_label_format)
            worksheet.set_row(row_num, 22)
        
        # Freeze panes
        worksheet.freeze_panes(9, 0)
        
        # ═══════════════ COMPANY WISE SHEET ═══════════════
        company_sheet = workbook.add_worksheet('Company Wise')
        
        company_sheet.set_column('A:A', 5)
        company_sheet.set_column('B:B', 30)
        company_sheet.set_column('C:C', 12)
        company_sheet.set_column('D:D', 18)
        company_sheet.set_column('E:E', 12)
        company_sheet.set_column('F:F', 12)
        
        company_sheet.merge_range('A1:F1', 'COMPANY WISE BREAKDOWN', title_format)
        company_sheet.set_row(0, 30)
        
        company_sheet.write('A3', 'SN', header_format)
        company_sheet.write('B3', 'Company Name', header_format)
        company_sheet.write('C3', 'Total', header_format)
        company_sheet.write('D3', 'Total Amount (NPR)', header_format)
        company_sheet.write('E3', 'Deposited', header_format)
        company_sheet.write('F3', 'Rejected', header_format)
        
        all_company_stats = cheques.values('company_name').annotate(
            total_count=Count('id'),
            total_amount=Sum('amount'),
            deposited_count=Count('id', filter=Q(status='Deposited')),
            rejected_count=Count('id', filter=Q(status='Rejected'))
        ).order_by('-total_amount')
        
        row = 4
        for idx, company in enumerate(all_company_stats, 1):
            company_sheet.write(f'A{row}', idx, date_format)
            company_sheet.write(f'B{row}', company['company_name'], cell_format)
            company_sheet.write(f'C{row}', company['total_count'], cell_format)
            company_sheet.write(f'D{row}', float(company['total_amount']) if company['total_amount'] else 0, number_format)
            company_sheet.write(f'E{row}', company['deposited_count'], cell_format)
            company_sheet.write(f'F{row}', company['rejected_count'], cell_format)
            row += 1
    
    # Prepare response
    output.seek(0)
    
    filename = f'Cheque_Report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    
    response = HttpResponse(
        output.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    
    return response


@login_required(login_url='/auth/login/')
def get_company_suggestions(request):
    """API endpoint for company name autocomplete"""
    query = request.GET.get('q', '')
    
    companies = ChequeModel.objects.filter(
        status__in=['Deposited', 'Rejected'],
        company_name__icontains=query
    ).values_list('company_name', flat=True).distinct()[:10]
    
    return JsonResponse(list(companies), safe=False)
#!==========================================================================================




# !================================= EMAIL DEPOSITED CHEQUES =============================================================
@login_required(login_url='/auth/login/')
def email_deposited_cheques(request):
    """Page to select deposited cheques and send via email"""
    
    # Get filter parameters
    date_from = request.GET.get('date_from', '')
    date_to = request.GET.get('date_to', '')
    company_name = request.GET.get('company_name', '')
    cheque_no = request.GET.get('cheque_no', '')
    
    # Check if any filter is applied
    has_filters = any([date_from, date_to, company_name, cheque_no])
    
    # Base queryset - Only Deposited cheques
    cheques = ChequeModel.objects.filter(status='Deposited').select_related('deposit')
    
    # Apply filters
    if date_from:
        cheques = cheques.filter(deposit__deposited_at__date__gte=date_from)
    if date_to:
        cheques = cheques.filter(deposit__deposited_at__date__lte=date_to)
    if company_name:
        cheques = cheques.filter(company_name__icontains=company_name)
    if cheque_no:
        cheques = cheques.filter(cheque_no__icontains=cheque_no)
    
    # Order by deposited date (most recent first)
    cheques = cheques.order_by('-deposit__deposited_at')
    
    # Calculate totals
    total_count = cheques.count()
    total_amount = cheques.aggregate(total=Sum('amount'))['total'] or 0
    
    # Get unique company names for dropdown
    all_companies = list(
        ChequeModel.objects.filter(status='Deposited')
        .values_list('company_name', flat=True)
        .distinct()
        .order_by('company_name')
    )
    
    # Current date for default date_to
    current_date = date.today().strftime('%Y-%m-%d')
    
    # Predefined email addresses
    email_options = [
        {'value': 'pioneersoftware@outlook.com', 'label': 'Pioneer Software (Outlook)'},
        {'value': 'info@pioneersoftware.com', 'label': 'Pioneer Software (Info)'},
    ]
    
    # Only show data if filters are applied
    if has_filters:
        # Pagination
        paginator = Paginator(cheques, 20)
        page_number = request.GET.get('page', 1)
        page_obj = paginator.get_page(page_number)
        
        current = page_obj.number
        total_pages = paginator.num_pages
        start_p = max(current - 2, 1)
        end_p = min(current + 2, total_pages)
        page_range = range(start_p, end_p + 1)
        
        start_index = (page_obj.number - 1) * paginator.per_page
    else:
        page_obj = None
        page_range = []
        start_index = 0
        total_count = 0
        total_amount = 0
    
    context = {
        'page_obj': page_obj,
        'page_range': page_range,
        'start_index': start_index,
        'total_count': total_count,
        'total_amount': total_amount,
        'email_options': email_options,
        'all_companies_json': json.dumps(all_companies),
        'has_filters': has_filters,
        'current_date': current_date,
        # Filter values for form persistence
        'filter_date_from': date_from,
        'filter_date_to': date_to if date_to else current_date,
        'filter_company': company_name,
        'filter_cheque_no': cheque_no,
    }
    
    return render(request, 'Report/email_deposited.html', context)


# !================================= EMAIL DEPOSITED CHEQUES =============================================================
@login_required(login_url='/auth/login/')
def send_deposited_email(request):
    """API endpoint to send email with selected deposited cheques"""
    
    if request.method != 'POST':
        return JsonResponse({'success': False, 'message': 'Invalid request method.'})
    
    try:
        data = json.loads(request.body)
        recipient_email = data.get('email', '').strip()
        cheque_ids = data.get('cheque_ids', [])
        
        # Validate
        if not recipient_email:
            return JsonResponse({'success': False, 'message': 'Please enter recipient email.'})
        
        if not cheque_ids:
            return JsonResponse({'success': False, 'message': 'Please select at least one cheque.'})
        
        # Get selected cheques
        cheques = ChequeModel.objects.filter(
            id__in=cheque_ids, 
            status='Deposited'
        ).select_related('deposit').order_by('-deposit__deposited_at')
        
        if not cheques.exists():
            return JsonResponse({'success': False, 'message': 'No valid cheques found.'})
        
        # Prepare data for Excel
        excel_data = []
        for cheque in cheques:
            row = {
                'Company Name': cheque.company_name,
                'Cheque No': cheque.cheque_no,
                'Amount (NPR)': float(cheque.amount),
                'Cheque Date': cheque.cheque_date.strftime('%Y-%m-%d') if cheque.cheque_date else '',
                'Deposited Date': cheque.deposit.deposited_at.strftime('%Y-%m-%d %H:%M') if cheque.deposit and cheque.deposit.deposited_at else '',
                'Bank Name': cheque.deposit.bank_name if cheque.deposit else '',
                'Branch': cheque.deposit.branch_name if cheque.deposit else '',
            }
            excel_data.append(row)
        
        # Calculate totals
        total_amount = sum(item['Amount (NPR)'] for item in excel_data)
        
        # Create Excel file
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Deposited Cheques')
            
            # Formats
            title_format = workbook.add_format({
                'bold': True, 'font_size': 16, 'font_color': '#1e1e2f',
                'align': 'center', 'valign': 'vcenter'
            })
            subtitle_format = workbook.add_format({
                'font_size': 11, 'font_color': '#6b7280',
                'align': 'center', 'valign': 'vcenter'
            })
            header_format = workbook.add_format({
                'bold': True, 'font_size': 11, 'font_color': 'white',
                'bg_color': '#1e1e2f', 'border': 1, 'align': 'center', 'valign': 'vcenter'
            })
            cell_format = workbook.add_format({
                'font_size': 10, 'border': 1, 'align': 'left', 'valign': 'vcenter'
            })
            date_format = workbook.add_format({
                'font_size': 10, 'border': 1, 'align': 'center', 'valign': 'vcenter'
            })
            number_format = workbook.add_format({
                'font_size': 10, 'border': 1, 'align': 'right',
                'valign': 'vcenter', 'num_format': 'NPR #,##0.00'
            })
            total_label_format = workbook.add_format({
                'bold': True, 'font_size': 11, 'font_color': 'white',
                'bg_color': '#16a34a', 'border': 1, 'align': 'right', 'valign': 'vcenter'
            })
            total_value_format = workbook.add_format({
                'bold': True, 'font_size': 11, 'font_color': 'white',
                'bg_color': '#16a34a', 'border': 1, 'align': 'right',
                'valign': 'vcenter', 'num_format': 'NPR #,##0.00'
            })
            
            # Column widths
            worksheet.set_column('A:A', 28)  # Company
            worksheet.set_column('B:B', 15)  # Cheque No
            worksheet.set_column('C:C', 18)  # Amount
            worksheet.set_column('D:D', 14)  # Cheque Date
            worksheet.set_column('E:E', 18)  # Deposited Date
            worksheet.set_column('F:F', 22)  # Bank
            worksheet.set_column('G:G', 18)  # Branch
            
            # Title
            worksheet.merge_range('A1:G1', 'DEPOSITED CHEQUES REPORT', title_format)
            worksheet.set_row(0, 28)
            
            # Subtitle
            worksheet.merge_range('A2:G2', f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M")} | Total Cheques: {len(excel_data)}', subtitle_format)
            worksheet.set_row(1, 20)
            
            # Headers
            headers = ['Company Name', 'Cheque No', 'Amount (NPR)', 'Cheque Date', 
                      'Deposited Date', 'Bank Name', 'Branch']
            for col, header in enumerate(headers):
                worksheet.write(3, col, header, header_format)
            worksheet.set_row(3, 22)
            
            # Data rows
            row_num = 4
            for item in excel_data:
                worksheet.write(row_num, 0, item['Company Name'], cell_format)
                worksheet.write(row_num, 1, item['Cheque No'], cell_format)
                worksheet.write(row_num, 2, item['Amount (NPR)'], number_format)
                worksheet.write(row_num, 3, item['Cheque Date'], date_format)
                worksheet.write(row_num, 4, item['Deposited Date'], date_format)
                worksheet.write(row_num, 5, item['Bank Name'], cell_format)
                worksheet.write(row_num, 6, item['Branch'], cell_format)
                worksheet.set_row(row_num, 18)
                row_num += 1
            
            # Total row
            worksheet.write(row_num, 0, '', total_label_format)
            worksheet.write(row_num, 1, 'TOTAL:', total_label_format)
            worksheet.write(row_num, 2, total_amount, total_value_format)
            worksheet.write(row_num, 3, '', total_label_format)
            worksheet.write(row_num, 4, '', total_label_format)
            worksheet.write(row_num, 5, '', total_label_format)
            worksheet.write(row_num, 6, '', total_label_format)
            worksheet.set_row(row_num, 22)
            
            # Freeze header
            worksheet.freeze_panes(4, 0)
        
        output.seek(0)
        
        # Prepare email
        subject = f'Deposited Cheques Report - {datetime.now().strftime("%Y-%m-%d")}'
        
        body = f"""
Dear Sir/Madam,

Please find attached the deposited cheques report.

Summary:
- Total Cheques: {len(excel_data)}
- Total Amount: NPR {total_amount:,.2f}
- Generated On: {datetime.now().strftime("%Y-%m-%d %H:%M")}

This is an auto-generated email from the Cheque Management System.

Best Regards,
{request.user.first_name} {request.user.last_name}
Cheque Management System
        """
        
        # Create email
        email = EmailMessage(
            subject=subject,
            body=body,
            from_email=settings.DEFAULT_FROM_EMAIL,
            to=[recipient_email],
        )
        
        # Attach Excel file
        filename = f'Deposited_Cheques_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        email.attach(
            filename,
            output.read(),
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        # Send email
        email.send(fail_silently=False)
        
        return JsonResponse({
            'success': True, 
            'message': f'Email sent successfully to {recipient_email}!'
        })
        
    except Exception as e:
        return JsonResponse({'success': False, 'message': str(e)})


# !============================ API to get selected cheques data for preview before emailing 
@login_required(login_url='/auth/login/')
def get_selected_cheques_data(request):
    """API endpoint to get data of selected cheques for preview"""
    
    if request.method != 'POST':
        return JsonResponse({'success': False, 'message': 'Invalid request method.'})
    
    try:
        data = json.loads(request.body)
        cheque_ids = data.get('cheque_ids', [])
        
        if not cheque_ids:
            return JsonResponse({'success': False, 'message': 'No cheques selected.'})
        
        cheques = ChequeModel.objects.filter(
            id__in=cheque_ids, 
            status='Deposited'
        ).select_related('deposit').order_by('-deposit__deposited_at')
        
        cheque_data = []
        total_amount = 0
        
        for cheque in cheques:
            total_amount += float(cheque.amount)
            cheque_data.append({
                'id': cheque.id,
                'company_name': cheque.company_name,
                'cheque_no': cheque.cheque_no,
                'amount': float(cheque.amount),
                'cheque_date': cheque.cheque_date.strftime('%Y-%m-%d') if cheque.cheque_date else '',
                'bank_name': cheque.deposit.bank_name if cheque.deposit else '',
                'branch_name': cheque.deposit.branch_name if cheque.deposit else '',
                'deposited_at': cheque.deposit.deposited_at.strftime('%Y-%m-%d %H:%M') if cheque.deposit and cheque.deposit.deposited_at else '',
            })
        
        return JsonResponse({
            'success': True,
            'cheques': cheque_data,
            'total_count': len(cheque_data),
            'total_amount': total_amount,
        })
        
    except Exception as e:
        return JsonResponse({'success': False, 'message': str(e)})
#!===================================== END OF DEPOSITED CHEQUES ===================================================== 
# ?===================================================================================================================



# !=========================================== KHAJA VIEWS ===========================================!


#!================================= KHAJA MAIN VIEW ===========================================! 
@login_required(login_url='/auth/login/')
def khaja_list(request):
    """Main Khaja page with date filter and pagination"""
    
    # Get filter parameters
    date_from = request.GET.get('date_from', '')
    date_to = request.GET.get('date_to', '')
    
    # Default to current date if not provided
    current_date = date.today().strftime('%Y-%m-%d')
    if not date_from:
        date_from = current_date
    if not date_to:
        date_to = current_date
    
    # Base queryset
    bills = KhajaBill.objects.prefetch_related('items')
    
    # Apply date filters
    if date_from:
        bills = bills.filter(date__gte=date_from)
    if date_to:
        bills = bills.filter(date__lte=date_to)
    
    # Order by date descending
    bills = bills.order_by('-date', '-created_at')
    
    # Calculate totals for filtered data
    total_count = bills.count()
    total_amount = sum(bill.total_amount for bill in bills)
    
    # Pagination
    paginator = Paginator(bills, 15)
    page_number = request.GET.get('page', 1)
    page_obj = paginator.get_page(page_number)
    
    current = page_obj.number
    total_pages = paginator.num_pages
    start_p = max(current - 2, 1)
    end_p = min(current + 2, total_pages)
    page_range = range(start_p, end_p + 1)
    
    start_index = (page_obj.number - 1) * paginator.per_page
    
    context = {
        'page_obj': page_obj,
        'page_range': page_range,
        'start_index': start_index,
        'total_count': total_count,
        'total_amount': total_amount,
        'current_date': current_date,
        'filter_date_from': date_from,
        'filter_date_to': date_to,
    }
    
    return render(request, 'Khaja/khaja.html', context)


# !================================= KHAJA PRINT VIEW ===========================================!
@login_required(login_url='/auth/login/')
def khaja_print(request, pk):
    """Print view for Khaja bill - thermal receipt format"""
    try:
        bill = KhajaBill.objects.prefetch_related('items').get(pk=pk)
        items = bill.items.all()
        
        # Convert amount to words
        def number_to_words(num):
            """Convert number to words for Nepali Rupees"""
            ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine',
                    'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen',
                    'Seventeen', 'Eighteen', 'Nineteen']
            tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']
            
            if num == 0:
                return 'Zero'
            
            def words(n):
                if n < 20:
                    return ones[n]
                elif n < 100:
                    return tens[n // 10] + ('' if n % 10 == 0 else ' ' + ones[n % 10])
                elif n < 1000:
                    return ones[n // 100] + ' Hundred' + ('' if n % 100 == 0 else ' ' + words(n % 100))
                elif n < 100000:
                    return words(n // 1000) + ' Thousand' + ('' if n % 1000 == 0 else ' ' + words(n % 1000))
                elif n < 10000000:
                    return words(n // 100000) + ' Lakh' + ('' if n % 100000 == 0 else ' ' + words(n % 100000))
                else:
                    return words(n // 10000000) + ' Crore' + ('' if n % 10000000 == 0 else ' ' + words(n % 10000000))
            
            rupees = int(num)
            paisa = int(round((num - rupees) * 100))
            
            result = 'Rupees ' + words(rupees)
            if paisa > 0:
                result += ' and ' + words(paisa) + ' Paisa'
            result += ' Only'
            
            return result
        
        amount_in_words = number_to_words(float(bill.total_amount))
        
        context = {
            'bill': bill,
            'items': items,
            'amount_in_words': amount_in_words,
        }
        
        return render(request, 'Khaja/khaja_print.html', context)
    
    except KhajaBill.DoesNotExist:
        from django.http import Http404
        raise Http404("Bill not found")

#!================================= KHAJA CRUD API VIEWS ===========================================! 
@login_required(login_url='/auth/login/')
def khaja_create(request):
    """Create a new Khaja bill with items"""
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            bill_date = data.get('date')
            items = data.get('items', [])
            
            if not bill_date:
                return JsonResponse({'success': False, 'message': 'Date is required.'})
            
            if not items or len(items) == 0:
                return JsonResponse({'success': False, 'message': 'At least one item is required.'})
            
            # Create bill
            bill = KhajaBill.objects.create(date=bill_date)
            
            # Create items
            for item in items:
                qty = Decimal(str(item.get('qty', 0)))
                rate = Decimal(str(item.get('rate', 0)))
                amount = qty * rate
                
                KhajaBillItem.objects.create(
                    bill=bill,
                    particular=item.get('particular', ''),
                    qty=qty,
                    rate=rate,
                    amount=amount
                )
            
            return JsonResponse({'success': True, 'message': 'Bill created successfully.'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    
    return JsonResponse({'success': False, 'message': 'Invalid request.'})


# !================================= KHAJA UPDATE API VIEW ===========================================!
@login_required(login_url='/auth/login/')
def khaja_update(request, pk):
    """Update an existing Khaja bill"""
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            bill = KhajaBill.objects.get(pk=pk)
            
            bill_date = data.get('date')
            items = data.get('items', [])
            
            if not bill_date:
                return JsonResponse({'success': False, 'message': 'Date is required.'})
            
            if not items or len(items) == 0:
                return JsonResponse({'success': False, 'message': 'At least one item is required.'})
            
            # Update bill date
            bill.date = bill_date
            bill.save()
            
            # Delete existing items and recreate
            bill.items.all().delete()
            
            for item in items:
                qty = Decimal(str(item.get('qty', 0)))
                rate = Decimal(str(item.get('rate', 0)))
                amount = qty * rate
                
                KhajaBillItem.objects.create(
                    bill=bill,
                    particular=item.get('particular', ''),
                    qty=qty,
                    rate=rate,
                    amount=amount
                )
            
            return JsonResponse({'success': True, 'message': 'Bill updated successfully.'})
        except KhajaBill.DoesNotExist:
            return JsonResponse({'success': False, 'message': 'Bill not found.'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    
    return JsonResponse({'success': False, 'message': 'Invalid request.'})


# !================================= KHAJA DELETE API VIEW ===========================================!
@login_required(login_url='/auth/login/')
def khaja_delete(request, pk):
    """Delete a Khaja bill"""
    if request.method == 'POST':
        try:
            bill = KhajaBill.objects.get(pk=pk)
            bill.delete()
            return JsonResponse({'success': True, 'message': 'Bill deleted successfully.'})
        except KhajaBill.DoesNotExist:
            return JsonResponse({'success': False, 'message': 'Bill not found.'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    
    return JsonResponse({'success': False, 'message': 'Invalid request.'})


# !================================= KHAJA DETAIL API VIEW ===========================================!
@login_required(login_url='/auth/login/')
def khaja_detail(request, pk):
    """Get bill details for view/edit modal"""
    try:
        bill = KhajaBill.objects.prefetch_related('items').get(pk=pk)
        
        items_data = []
        for item in bill.items.all():
            items_data.append({
                'id': item.id,
                'particular': item.particular,
                'qty': float(item.qty),
                'rate': float(item.rate),
                'amount': float(item.amount),
            })
        
        data = {
            'id': bill.id,
            'date': bill.date.strftime('%Y-%m-%d'),
            'items': items_data,
            'total_amount': float(bill.total_amount),
            'created_at': bill.created_at.strftime('%Y-%m-%d %H:%M'),
        }
        
        return JsonResponse({'success': True, 'data': data})
    except KhajaBill.DoesNotExist:
        return JsonResponse({'success': False, 'message': 'Bill not found.'})
    except Exception as e:
        return JsonResponse({'success': False, 'message': str(e)})


# !================================= KHAJA EXPORT EXCEL API VIEW ===========================================!
@login_required(login_url='/auth/login/')
def khaja_export_excel(request):
    """Export filtered Khaja data to Excel"""
    
    # Get filter parameters
    date_from = request.GET.get('date_from', '')
    date_to = request.GET.get('date_to', '')
    
    # Default to current date
    current_date = date.today().strftime('%Y-%m-%d')
    if not date_from:
        date_from = current_date
    if not date_to:
        date_to = current_date
    
    # Base queryset
    bills = KhajaBill.objects.prefetch_related('items')
    
    if date_from:
        bills = bills.filter(date__gte=date_from)
    if date_to:
        bills = bills.filter(date__lte=date_to)
    
    bills = bills.order_by('-date', '-created_at')
    
    # Create Excel
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formats
        title_format = workbook.add_format({
            'bold': True, 'font_size': 18, 'font_color': '#1e1e2f',
            'align': 'center', 'valign': 'vcenter'
        })
        header_format = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_color': 'white',
            'bg_color': '#4f46e5', 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        cell_format = workbook.add_format({
            'font_size': 10, 'border': 1, 'align': 'left', 'valign': 'vcenter'
        })
        number_format = workbook.add_format({
            'font_size': 10, 'border': 1, 'align': 'right', 'valign': 'vcenter',
            'num_format': '₹#,##0.00'
        })
        date_format = workbook.add_format({
            'font_size': 10, 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        total_format = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_color': 'white',
            'bg_color': '#1e1e2f', 'border': 1, 'align': 'right', 'valign': 'vcenter',
            'num_format': '₹#,##0.00'
        })
        
        # Summary Sheet
        summary_sheet = workbook.add_worksheet('Summary')
        summary_sheet.set_column('A:A', 8)
        summary_sheet.set_column('B:B', 15)
        summary_sheet.set_column('C:C', 40)
        summary_sheet.set_column('D:D', 18)
        
        summary_sheet.merge_range('A1:D1', 'KHAJA BILL SUMMARY', title_format)
        summary_sheet.set_row(0, 30)
        
        summary_sheet.merge_range('A2:D2', f'Date Range: {date_from} to {date_to}', workbook.add_format({
            'font_size': 11, 'align': 'center', 'font_color': '#6b7280'
        }))
        
        summary_sheet.write('A4', 'SN', header_format)
        summary_sheet.write('B4', 'Date', header_format)
        summary_sheet.write('C4', 'Particulars', header_format)
        summary_sheet.write('D4', 'Amount', header_format)
        
        row = 5
        grand_total = 0
        for idx, bill in enumerate(bills, 1):
            particulars = ', '.join([item.particular for item in bill.items.all()])
            total = float(bill.total_amount)
            grand_total += total
            
            summary_sheet.write(f'A{row}', idx, date_format)
            summary_sheet.write(f'B{row}', bill.date.strftime('%Y-%m-%d'), date_format)
            summary_sheet.write(f'C{row}', particulars, cell_format)
            summary_sheet.write(f'D{row}', total, number_format)
            row += 1
        
        # Total row
        summary_sheet.write(f'A{row}', '', total_format)
        summary_sheet.write(f'B{row}', '', total_format)
        summary_sheet.write(f'C{row}', 'GRAND TOTAL', total_format)
        summary_sheet.write(f'D{row}', grand_total, total_format)
        
        # Detailed Sheet
        detail_sheet = workbook.add_worksheet('Detailed')
        detail_sheet.set_column('A:A', 8)
        detail_sheet.set_column('B:B', 12)
        detail_sheet.set_column('C:C', 8)
        detail_sheet.set_column('D:D', 30)
        detail_sheet.set_column('E:E', 10)
        detail_sheet.set_column('F:F', 12)
        detail_sheet.set_column('G:G', 15)
        
        detail_sheet.merge_range('A1:G1', 'KHAJA BILL DETAILS', title_format)
        detail_sheet.set_row(0, 30)
        
        detail_sheet.write('A3', 'Bill #', header_format)
        detail_sheet.write('B3', 'Date', header_format)
        detail_sheet.write('C3', 'Item #', header_format)
        detail_sheet.write('D3', 'Particular', header_format)
        detail_sheet.write('E3', 'Qty', header_format)
        detail_sheet.write('F3', 'Rate', header_format)
        detail_sheet.write('G3', 'Amount', header_format)
        
        row = 4
        for bill in bills:
            item_num = 1
            for item in bill.items.all():
                detail_sheet.write(f'A{row}', bill.id, date_format)
                detail_sheet.write(f'B{row}', bill.date.strftime('%Y-%m-%d'), date_format)
                detail_sheet.write(f'C{row}', item_num, date_format)
                detail_sheet.write(f'D{row}', item.particular, cell_format)
                detail_sheet.write(f'E{row}', float(item.qty), number_format)
                detail_sheet.write(f'F{row}', float(item.rate), number_format)
                detail_sheet.write(f'G{row}', float(item.amount), number_format)
                row += 1
                item_num += 1
    
    output.seek(0)
    filename = f'Khaja_Report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    
    response = HttpResponse(
        output.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    
    return response
#!=========================================== END OF KHAJA VIEWS ===========================================!
# ?=================================================================================================================== 