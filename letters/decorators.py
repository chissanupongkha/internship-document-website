from functools import wraps

from django.contrib import messages
from django.shortcuts import redirect


def get_role(user):
    """admin / staff / viewer. Accounts made before roles existed keep working:
    superusers count as admin, everyone else as staff."""
    if not user.is_authenticated:
        return None
    profile = getattr(user, 'profile', None) if hasattr(user, 'profile') else None
    if profile is not None:
        return profile.role
    return 'admin' if user.is_superuser else 'staff'


def _blocked(request, text):
    messages.error(request, text)
    return redirect('letters:login')


def staff_required(view_func):
    """Logged-in, active, not-expired staff. Viewers may only look (GET)."""
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        user = request.user
        if not user.is_authenticated:
            return redirect('letters:login')
        if not user.is_staff:
            return _blocked(request, "Access restricted to authorized staff members only.")
        profile = getattr(user, 'profile', None) if hasattr(user, 'profile') else None
        if profile is not None and profile.expired:
            from django.contrib.auth import logout
            logout(request)
            return _blocked(request, "This account has expired. Ask an admin to extend it.")
        if get_role(user) == 'viewer' and request.method not in ('GET', 'HEAD'):
            messages.error(request, "Your role can look but not change anything.")
            return redirect(request.META.get('HTTP_REFERER') or 'letters:index')
        return view_func(request, *args, **kwargs)
    return _wrapped_view


def admin_required(view_func):
    """Admin role only (accounts, configuration)."""
    @wraps(view_func)
    @staff_required
    def _wrapped_view(request, *args, **kwargs):
        if get_role(request.user) != 'admin':
            messages.error(request, "Only an admin can open that page.")
            return redirect('letters:index')
        return view_func(request, *args, **kwargs)
    return _wrapped_view
