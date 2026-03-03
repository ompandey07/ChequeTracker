import json
from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.models import User
from django.http import JsonResponse


def ensure_default_user():
    if not User.objects.filter(email='manimaharjan@admin.com').exists():
        user = User.objects.create_user(
            username='manimaharjan',
            email='manimaharjan@admin.com',
            password='mani@1200',
            first_name='Mani',
            last_name='Maharjan',
        )
        user.is_staff = True
        user.save()


def login_view(request):
    ensure_default_user()

    if request.user.is_authenticated:
        return redirect('admin_dashboard')

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except json.JSONDecodeError:
            return JsonResponse({'success': False, 'message': 'Invalid request.'})

        email = data.get('email', '').strip()
        password = data.get('password', '')

        if not email or not password:
            return JsonResponse({'success': False, 'message': 'Email and password are required.'})

        try:
            user_obj = User.objects.get(email=email)
        except User.DoesNotExist:
            return JsonResponse({'success': False, 'message': 'Invalid email or password.'})

        user = authenticate(request, username=user_obj.username, password=password)

        if user is not None:
            login(request, user)
            return JsonResponse({'success': True, 'redirect': '/adminview/admin_dashboard/'})
        else:
            return JsonResponse({'success': False, 'message': 'Invalid email or password.'})

    return render(request, 'Auth/login.html')


def logout_view(request):
    logout(request)
    return redirect('login')