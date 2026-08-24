from functools import wraps
from django.shortcuts import redirect
from django.contrib import messages

def staff_required(view_func):
    """Restricts access to authenticated staff members only."""
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect('letters:login')
        if not request.user.is_staff:
            messages.error(request, "Access restricted to authorized staff members only.")
            return redirect('letters:login')
        return view_func(request, *args, **kwargs)
    return _wrapped_view