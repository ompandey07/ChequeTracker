from django.db import models


class ChequeModel(models.Model):
    STATUS_CHOICES = [
        ('New', 'New'),
        ('Deposited', 'Deposited'),
        ('Rejected', 'Rejected'),
    ]
    
    company_name = models.CharField(max_length=255)
    cheque_no = models.CharField(max_length=100)
    amount = models.DecimalField(max_digits=12, decimal_places=2)
    cheque_date = models.DateField()
    remarks = models.TextField(blank=True, null=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='New')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f"{self.company_name} — {self.cheque_no}"


class DepositModel(models.Model):
    cheque = models.OneToOneField(
        ChequeModel, 
        on_delete=models.CASCADE, 
        related_name='deposit'
    )
    bank_name = models.CharField(max_length=255)
    branch_name = models.CharField(max_length=255)
    deposit_slip = models.ImageField(upload_to='deposit_slips/')
    deposited_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Deposit for {self.cheque.cheque_no} at {self.bank_name}"
    

class KhajaBill(models.Model):
    date = models.DateField()
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-date', '-created_at']

    def __str__(self):
        return f"Bill #{self.id} - {self.date}"
    
    @property
    def total_amount(self):
        return self.items.aggregate(total=models.Sum('amount'))['total'] or 0
    
    @property
    def item_count(self):
        return self.items.count()


class KhajaBillItem(models.Model):
    bill = models.ForeignKey(KhajaBill, on_delete=models.CASCADE, related_name='items')
    particular = models.CharField(max_length=255)
    qty = models.DecimalField(max_digits=10, decimal_places=2)
    rate = models.DecimalField(max_digits=12, decimal_places=2)
    amount = models.DecimalField(max_digits=12, decimal_places=2)

    class Meta:
        ordering = ['id']

    def __str__(self):
        return f"{self.particular} - {self.qty} x {self.rate}"
    
    def save(self, *args, **kwargs):
        self.amount = self.qty * self.rate
        super().save(*args, **kwargs)