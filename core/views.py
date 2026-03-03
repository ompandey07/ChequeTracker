from django.db.models.functions import TruncMonth, TruncDate
from django.contrib.auth.decorators import login_required
from django.views.decorators.http import require_POST
from django.http import JsonResponse, HttpResponse
from django.shortcuts import render, redirect
from .models import ChequeModel, DepositModel
from django.core.paginator import Paginator
from django.db.models import Sum, Count, Q
from django.core.mail import EmailMessage
from django.conf import settings
from datetime import datetime, timedelta
from django.db.models import Sum, Count
from django.http import JsonResponse
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
    """Main reports page with filters"""
    
    # Get filter parameters
    date_from = request.GET.get('date_from', '')
    date_to = request.GET.get('date_to', '')
    company_name = request.GET.get('company_name', '')
    status = request.GET.get('status', '')
    cheque_no = request.GET.get('cheque_no', '')
    
    # Base queryset
    cheques = ChequeModel.objects.all()
    
    # Apply filters
    if date_from:
        cheques = cheques.filter(cheque_date__gte=date_from)
    if date_to:
        cheques = cheques.filter(cheque_date__lte=date_to)
    if company_name:
        cheques = cheques.filter(company_name__icontains=company_name)
    if status:
        cheques = cheques.filter(status=status)
    if cheque_no:
        cheques = cheques.filter(cheque_no__icontains=cheque_no)
    
    # Calculate statistics for filtered data
    total_count = cheques.count()
    total_amount = cheques.aggregate(total=Sum('amount'))['total'] or 0
    
    # Status breakdown
    status_breakdown = cheques.values('status').annotate(
        count=Count('id'),
        total=Sum('amount')
    )
    
    new_count = 0
    new_amount = 0
    deposited_count = 0
    deposited_amount = 0
    rejected_count = 0
    rejected_amount = 0
    
    for item in status_breakdown:
        if item['status'] == 'New':
            new_count = item['count']
            new_amount = item['total'] or 0
        elif item['status'] == 'Deposited':
            deposited_count = item['count']
            deposited_amount = item['total'] or 0
        elif item['status'] == 'Rejected':
            rejected_count = item['count']
            rejected_amount = item['total'] or 0
    
    # Monthly breakdown for chart
    monthly_data = cheques.annotate(
        month=TruncMonth('cheque_date')
    ).values('month').annotate(
        count=Count('id'),
        total=Sum('amount')
    ).order_by('month')
    
    months = []
    monthly_counts = []
    monthly_amounts = []
    for item in monthly_data:
        if item['month']:
            months.append(item['month'].strftime('%b %Y'))
            monthly_counts.append(item['count'])
            monthly_amounts.append(float(item['total']) if item['total'] else 0)
    
    # Company breakdown
    company_data = cheques.values('company_name').annotate(
        count=Count('id'),
        total=Sum('amount')
    ).order_by('-total')[:10]
    
    company_names = [item['company_name'][:15] for item in company_data]
    company_amounts = [float(item['total']) if item['total'] else 0 for item in company_data]
    
    # Get unique company names for dropdown
    all_companies = ChequeModel.objects.values_list('company_name', flat=True).distinct().order_by('company_name')
    
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
    
    context = {
        'page_obj': page_obj,
        'page_range': page_range,
        'start_index': start_index,
        'total_count': total_count,
        'total_amount': total_amount,
        'new_count': new_count,
        'new_amount': new_amount,
        'deposited_count': deposited_count,
        'deposited_amount': deposited_amount,
        'rejected_count': rejected_count,
        'rejected_amount': rejected_amount,
        'all_companies': all_companies,
        # Filter values for form persistence
        'filter_date_from': date_from,
        'filter_date_to': date_to,
        'filter_company': company_name,
        'filter_status': status,
        'filter_cheque_no': cheque_no,
        # Chart data
        'months_json': json.dumps(months),
        'monthly_counts_json': json.dumps(monthly_counts),
        'monthly_amounts_json': json.dumps(monthly_amounts),
        'company_names_json': json.dumps(company_names),
        'company_amounts_json': json.dumps(company_amounts),
    }
    
    return render(request, 'Report/reports.html', context)


@login_required(login_url='/auth/login/')
def export_report_excel(request):
    """Export filtered data to beautiful Excel report"""
    
    # Get filter parameters
    date_from = request.GET.get('date_from', '')
    date_to = request.GET.get('date_to', '')
    company_name = request.GET.get('company_name', '')
    status = request.GET.get('status', '')
    cheque_no = request.GET.get('cheque_no', '')
    
    # Base queryset
    cheques = ChequeModel.objects.all()
    
    # Apply filters
    if date_from:
        cheques = cheques.filter(cheque_date__gte=date_from)
    if date_to:
        cheques = cheques.filter(cheque_date__lte=date_to)
    if company_name:
        cheques = cheques.filter(company_name__icontains=company_name)
    if status:
        cheques = cheques.filter(status=status)
    if cheque_no:
        cheques = cheques.filter(cheque_no__icontains=cheque_no)
    
    # Prepare data for Excel
    data = []
    for idx, cheque in enumerate(cheques, 1):
        row = {
            'SN': idx,
            'Company Name': cheque.company_name,
            'Cheque No': cheque.cheque_no,
            'Amount': float(cheque.amount),
            'Cheque Date': cheque.cheque_date.strftime('%Y-%m-%d') if cheque.cheque_date else '',
            'Status': cheque.status,
            'Remarks': cheque.remarks or '',
            'Created At': cheque.created_at.strftime('%Y-%m-%d %H:%M') if cheque.created_at else '',
        }
        
        # Add deposit info if deposited
        if cheque.status == 'Deposited' and hasattr(cheque, 'deposit'):
            row['Bank Name'] = cheque.deposit.bank_name
            row['Branch'] = cheque.deposit.branch_name
            row['Deposited At'] = cheque.deposit.deposited_at.strftime('%Y-%m-%d %H:%M') if cheque.deposit.deposited_at else ''
        else:
            row['Bank Name'] = ''
            row['Branch'] = ''
            row['Deposited At'] = ''
        
        data.append(row)
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Calculate summary statistics
    total_amount = cheques.aggregate(total=Sum('amount'))['total'] or 0
    new_amount = cheques.filter(status='New').aggregate(total=Sum('amount'))['total'] or 0
    deposited_amount = cheques.filter(status='Deposited').aggregate(total=Sum('amount'))['total'] or 0
    rejected_amount = cheques.filter(status='Rejected').aggregate(total=Sum('amount'))['total'] or 0
    
    # Create Excel file with formatting
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # ═══════════════ FORMATS ═══════════════
        # Header format
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
        
        # Title format
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'font_color': '#1e1e2f',
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Subtitle format
        subtitle_format = workbook.add_format({
            'font_size': 11,
            'font_color': '#6b7280',
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Cell format
        cell_format = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        # Number format
        number_format = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00'
        })
        
        # Date format
        date_format = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Status formats
        status_new_format = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#dbeafe',
            'font_color': '#1d4ed8',
            'bold': True
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
        
        # Summary label format
        summary_label_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': '#374151',
            'bg_color': '#f3f4f6',
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        # Summary value format
        summary_value_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': '#1e1e2f',
            'bg_color': '#f3f4f6',
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '₹#,##0.00'
        })
        
        # Total row format
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
            'num_format': '₹#,##0.00'
        })
        
        # ═══════════════ MAIN REPORT SHEET ═══════════════
        worksheet = workbook.add_worksheet('Cheque Report')
        
        # Set column widths
        worksheet.set_column('A:A', 6)   # SN
        worksheet.set_column('B:B', 25)  # Company Name
        worksheet.set_column('C:C', 15)  # Cheque No
        worksheet.set_column('D:D', 15)  # Amount
        worksheet.set_column('E:E', 12)  # Cheque Date
        worksheet.set_column('F:F', 12)  # Status
        worksheet.set_column('G:G', 25)  # Remarks
        worksheet.set_column('H:H', 16)  # Created At
        worksheet.set_column('I:I', 20)  # Bank Name
        worksheet.set_column('J:J', 15)  # Branch
        worksheet.set_column('K:K', 16)  # Deposited At
        
        # Title
        worksheet.merge_range('A1:K1', 'CHEQUE MANAGEMENT REPORT', title_format)
        worksheet.set_row(0, 30)
        
        # Subtitle with date range
        filter_text = f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        if date_from or date_to:
            filter_text += f' | Date Range: {date_from or "Start"} to {date_to or "End"}'
        if company_name:
            filter_text += f' | Company: {company_name}'
        if status:
            filter_text += f' | Status: {status}'
        
        worksheet.merge_range('A2:K2', filter_text, subtitle_format)
        worksheet.set_row(1, 20)
        
        # Empty row
        worksheet.set_row(2, 10)
        
        # Summary section
        worksheet.merge_range('A4:B4', 'SUMMARY', header_format)
        worksheet.merge_range('C4:D4', '', header_format)
        
        worksheet.write('A5', 'Total Cheques:', summary_label_format)
        worksheet.write('B5', len(data), workbook.add_format({
            'bold': True, 'font_size': 11, 'bg_color': '#f3f4f6', 'border': 1, 'align': 'center'
        }))
        worksheet.write('C5', 'Total Amount:', summary_label_format)
        worksheet.write('D5', float(total_amount), summary_value_format)
        
        worksheet.write('A6', 'New:', summary_label_format)
        worksheet.write('B6', cheques.filter(status='New').count(), workbook.add_format({
            'bold': True, 'font_size': 11, 'bg_color': '#dbeafe', 'border': 1, 'align': 'center', 'font_color': '#1d4ed8'
        }))
        worksheet.write('C6', 'New Amount:', summary_label_format)
        worksheet.write('D6', float(new_amount), summary_value_format)
        
        worksheet.write('A7', 'Deposited:', summary_label_format)
        worksheet.write('B7', cheques.filter(status='Deposited').count(), workbook.add_format({
            'bold': True, 'font_size': 11, 'bg_color': '#dcfce7', 'border': 1, 'align': 'center', 'font_color': '#16a34a'
        }))
        worksheet.write('C7', 'Deposited Amount:', summary_label_format)
        worksheet.write('D7', float(deposited_amount), summary_value_format)
        
        worksheet.write('A8', 'Rejected:', summary_label_format)
        worksheet.write('B8', cheques.filter(status='Rejected').count(), workbook.add_format({
            'bold': True, 'font_size': 11, 'bg_color': '#fee2e2', 'border': 1, 'align': 'center', 'font_color': '#dc2626'
        }))
        worksheet.write('C8', 'Rejected Amount:', summary_label_format)
        worksheet.write('D8', float(rejected_amount), summary_value_format)
        
        # Empty row
        worksheet.set_row(8, 15)
        
        # Data header row
        headers = ['SN', 'Company Name', 'Cheque No', 'Amount', 'Cheque Date', 
                   'Status', 'Remarks', 'Created At', 'Bank Name', 'Branch', 'Deposited At']
        
        for col, header in enumerate(headers):
            worksheet.write(9, col, header, header_format)
        worksheet.set_row(9, 25)
        
        # Data rows
        row_num = 10
        for item in data:
            worksheet.write(row_num, 0, item['SN'], date_format)
            worksheet.write(row_num, 1, item['Company Name'], cell_format)
            worksheet.write(row_num, 2, item['Cheque No'], cell_format)
            worksheet.write(row_num, 3, item['Amount'], number_format)
            worksheet.write(row_num, 4, item['Cheque Date'], date_format)
            
            # Status with conditional formatting
            status_val = item['Status']
            if status_val == 'New':
                worksheet.write(row_num, 5, status_val, status_new_format)
            elif status_val == 'Deposited':
                worksheet.write(row_num, 5, status_val, status_deposited_format)
            else:
                worksheet.write(row_num, 5, status_val, status_rejected_format)
            
            worksheet.write(row_num, 6, item['Remarks'], cell_format)
            worksheet.write(row_num, 7, item['Created At'], date_format)
            worksheet.write(row_num, 8, item['Bank Name'], cell_format)
            worksheet.write(row_num, 9, item['Branch'], cell_format)
            worksheet.write(row_num, 10, item['Deposited At'], date_format)
            
            worksheet.set_row(row_num, 20)
            row_num += 1
        
        # Total row
        if data:
            worksheet.write(row_num, 0, '', total_label_format)
            worksheet.write(row_num, 1, '', total_label_format)
            worksheet.write(row_num, 2, 'TOTAL:', total_label_format)
            worksheet.write(row_num, 3, float(total_amount), total_value_format)
            for col in range(4, 11):
                worksheet.write(row_num, col, '', total_label_format)
            worksheet.set_row(row_num, 25)
        
        # Freeze panes
        worksheet.freeze_panes(10, 0)
        
        # ═══════════════ SUMMARY SHEET ═══════════════
        summary_sheet = workbook.add_worksheet('Summary')
        
        summary_sheet.set_column('A:A', 25)
        summary_sheet.set_column('B:B', 15)
        summary_sheet.set_column('C:C', 20)
        
        summary_sheet.merge_range('A1:C1', 'CHEQUE SUMMARY REPORT', title_format)
        summary_sheet.set_row(0, 30)
        
        summary_sheet.merge_range('A2:C2', f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}', subtitle_format)
        
        # Summary table
        summary_sheet.write('A4', 'Category', header_format)
        summary_sheet.write('B4', 'Count', header_format)
        summary_sheet.write('C4', 'Amount', header_format)
        
        summary_sheet.write('A5', 'Total Cheques', summary_label_format)
        summary_sheet.write('B5', len(data), cell_format)
        summary_sheet.write('C5', float(total_amount), number_format)
        
        summary_sheet.write('A6', 'New / Pending', status_new_format)
        summary_sheet.write('B6', cheques.filter(status='New').count(), cell_format)
        summary_sheet.write('C6', float(new_amount), number_format)
        
        summary_sheet.write('A7', 'Deposited', status_deposited_format)
        summary_sheet.write('B7', cheques.filter(status='Deposited').count(), cell_format)
        summary_sheet.write('C7', float(deposited_amount), number_format)
        
        summary_sheet.write('A8', 'Rejected', status_rejected_format)
        summary_sheet.write('B8', cheques.filter(status='Rejected').count(), cell_format)
        summary_sheet.write('C8', float(rejected_amount), number_format)
        
        # Company breakdown
        company_breakdown = cheques.values('company_name').annotate(
            count=Count('id'),
            total=Sum('amount')
        ).order_by('-total')[:15]
        
        summary_sheet.merge_range('A11:C11', 'TOP COMPANIES BY AMOUNT', header_format)
        
        summary_sheet.write('A12', 'Company Name', header_format)
        summary_sheet.write('B12', 'Cheques', header_format)
        summary_sheet.write('C12', 'Amount', header_format)
        
        row = 13
        for company in company_breakdown:
            summary_sheet.write(f'A{row}', company['company_name'], cell_format)
            summary_sheet.write(f'B{row}', company['count'], cell_format)
            summary_sheet.write(f'C{row}', float(company['total']) if company['total'] else 0, number_format)
            row += 1
        
        # Add chart
        chart = workbook.add_chart({'type': 'pie'})
        chart.add_series({
            'name': 'Status Distribution',
            'categories': '=Summary!$A$5:$A$8',
            'values': '=Summary!$C$5:$C$8',
            'data_labels': {'percentage': True},
        })
        chart.set_title({'name': 'Amount by Status'})
        chart.set_style(10)
        summary_sheet.insert_chart('E4', chart, {'x_scale': 1.2, 'y_scale': 1.2})
        
        # ═══════════════ COMPANY WISE SHEET ═══════════════
        company_sheet = workbook.add_worksheet('Company Wise')
        
        company_sheet.set_column('A:A', 5)
        company_sheet.set_column('B:B', 30)
        company_sheet.set_column('C:C', 12)
        company_sheet.set_column('D:D', 18)
        company_sheet.set_column('E:E', 12)
        company_sheet.set_column('F:F', 12)
        company_sheet.set_column('G:G', 12)
        
        company_sheet.merge_range('A1:G1', 'COMPANY WISE BREAKDOWN', title_format)
        company_sheet.set_row(0, 30)
        
        company_sheet.write('A3', 'SN', header_format)
        company_sheet.write('B3', 'Company Name', header_format)
        company_sheet.write('C3', 'Total', header_format)
        company_sheet.write('D3', 'Total Amount', header_format)
        company_sheet.write('E3', 'New', header_format)
        company_sheet.write('F3', 'Deposited', header_format)
        company_sheet.write('G3', 'Rejected', header_format)
        
        all_company_stats = cheques.values('company_name').annotate(
            total_count=Count('id'),
            total_amount=Sum('amount'),
            new_count=Count('id', filter=Q(status='New')),
            deposited_count=Count('id', filter=Q(status='Deposited')),
            rejected_count=Count('id', filter=Q(status='Rejected'))
        ).order_by('-total_amount')
        
        row = 4
        for idx, company in enumerate(all_company_stats, 1):
            company_sheet.write(f'A{row}', idx, date_format)
            company_sheet.write(f'B{row}', company['company_name'], cell_format)
            company_sheet.write(f'C{row}', company['total_count'], cell_format)
            company_sheet.write(f'D{row}', float(company['total_amount']) if company['total_amount'] else 0, number_format)
            company_sheet.write(f'E{row}', company['new_count'], cell_format)
            company_sheet.write(f'F{row}', company['deposited_count'], cell_format)
            company_sheet.write(f'G{row}', company['rejected_count'], cell_format)
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
        company_name__icontains=query
    ).values_list('company_name', flat=True).distinct()[:10]
    
    return JsonResponse(list(companies), safe=False)

#!==========================================================================================



@login_required(login_url='/auth/login/')
def email_deposited_cheques(request):
    """Page to select deposited cheques and send via email"""
    
    cheques = ChequeModel.objects.filter(status='Deposited').select_related('deposit')
    
    # Calculate totals
    total_count = cheques.count()
    total_amount = cheques.aggregate(total=Sum('amount'))['total'] or 0
    
    # Pagination
    paginator = Paginator(cheques, 25)
    page_number = request.GET.get('page', 1)
    page_obj = paginator.get_page(page_number)
    
    current = page_obj.number
    total_pages = paginator.num_pages
    start_p = max(current - 2, 1)
    end_p = min(current + 2, total_pages)
    page_range = range(start_p, end_p + 1)
    
    start_index = (page_obj.number - 1) * paginator.per_page
    
    # Predefined email addresses
    email_options = [
        {'value': 'pioneersoftware@outlook.com', 'label': 'Pioneer Software (Outlook)'},
        {'value': 'info@pioneersoftware.com', 'label': 'Pioneer Software (Info)'},
    ]
    
    context = {
        'page_obj': page_obj,
        'page_range': page_range,
        'start_index': start_index,
        'total_count': total_count,
        'total_amount': total_amount,
        'email_options': email_options,
    }
    
    return render(request, 'Report/email_deposited.html', context)


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
        ).select_related('deposit')
        
        if not cheques.exists():
            return JsonResponse({'success': False, 'message': 'No valid cheques found.'})
        
        # Prepare data for Excel
        excel_data = []
        for cheque in cheques:
            row = {
                'Deposited At': cheque.deposit.deposited_at.strftime('%Y-%m-%d %H:%M') if cheque.deposit and cheque.deposit.deposited_at else '',
                'Deposited Bank': cheque.deposit.bank_name if cheque.deposit else '',
                'Deposited Branch': cheque.deposit.branch_name if cheque.deposit else '',
                'Company Name': cheque.company_name,
                'Cheque No': cheque.cheque_no,
                'Amount': float(cheque.amount),
                'Cheque Date': cheque.cheque_date.strftime('%Y-%m-%d') if cheque.cheque_date else '',
            }
            excel_data.append(row)
        
        # Calculate totals
        total_amount = sum(item['Amount'] for item in excel_data)
        
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
                'valign': 'vcenter', 'num_format': '#,##0.00'
            })
            total_label_format = workbook.add_format({
                'bold': True, 'font_size': 11, 'font_color': 'white',
                'bg_color': '#16a34a', 'border': 1, 'align': 'right', 'valign': 'vcenter'
            })
            total_value_format = workbook.add_format({
                'bold': True, 'font_size': 11, 'font_color': 'white',
                'bg_color': '#16a34a', 'border': 1, 'align': 'right',
                'valign': 'vcenter', 'num_format': '₹#,##0.00'
            })
            
            # Column widths
            worksheet.set_column('A:A', 18)  # Deposited At
            worksheet.set_column('B:B', 22)  # Bank
            worksheet.set_column('C:C', 18)  # Branch
            worksheet.set_column('D:D', 28)  # Company
            worksheet.set_column('E:E', 15)  # Cheque No
            worksheet.set_column('F:F', 15)  # Amount
            worksheet.set_column('G:G', 14)  # Cheque Date
            
            # Title
            worksheet.merge_range('A1:G1', 'DEPOSITED CHEQUES REPORT', title_format)
            worksheet.set_row(0, 28)
            
            # Subtitle
            worksheet.merge_range('A2:G2', f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M")} | Total Cheques: {len(excel_data)}', subtitle_format)
            worksheet.set_row(1, 20)
            
            # Headers
            headers = ['Deposited At', 'Deposited Bank', 'Deposited Branch', 
                      'Company Name', 'Cheque No', 'Amount', 'Cheque Date']
            for col, header in enumerate(headers):
                worksheet.write(3, col, header, header_format)
            worksheet.set_row(3, 22)
            
            # Data rows
            row_num = 4
            for item in excel_data:
                worksheet.write(row_num, 0, item['Deposited At'], date_format)
                worksheet.write(row_num, 1, item['Deposited Bank'], cell_format)
                worksheet.write(row_num, 2, item['Deposited Branch'], cell_format)
                worksheet.write(row_num, 3, item['Company Name'], cell_format)
                worksheet.write(row_num, 4, item['Cheque No'], cell_format)
                worksheet.write(row_num, 5, item['Amount'], number_format)
                worksheet.write(row_num, 6, item['Cheque Date'], date_format)
                worksheet.set_row(row_num, 18)
                row_num += 1
            
            # Total row
            worksheet.write(row_num, 0, '', total_label_format)
            worksheet.write(row_num, 1, '', total_label_format)
            worksheet.write(row_num, 2, '', total_label_format)
            worksheet.write(row_num, 3, '', total_label_format)
            worksheet.write(row_num, 4, 'TOTAL:', total_label_format)
            worksheet.write(row_num, 5, total_amount, total_value_format)
            worksheet.write(row_num, 6, '', total_label_format)
            worksheet.set_row(row_num, 22)
            
            # Freeze header
            worksheet.freeze_panes(4, 0)
        
        output.seek(0)
        
        # Prepare email
        subject = f'Deposited Cheques Data - {datetime.now().strftime("%Y-%m-%d")}'
        
        body = f"""
Dear Sir/Madam,

Please find attached the deposited cheques report.

Summary:
- Total Cheques: {len(excel_data)}
- Total Amount: ₹{total_amount:,.2f}
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
        ).select_related('deposit')
        
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