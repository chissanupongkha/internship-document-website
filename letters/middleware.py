from .models import AuditLog


class ActivityMiddleware:
    """Records every change a logged-in user makes (POST, PUT, DELETE) and every
    download or export, unless the view already wrote its own log line."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)
        try:
            user = getattr(request, 'user', None)
            if user is not None and user.is_authenticated and not getattr(request, '_activity_logged', False):
                path = request.path
                changed = request.method in ('POST', 'PUT', 'PATCH', 'DELETE')
                exported = request.method == 'GET' and ('download' in path or 'export' in path)
                if (changed or exported) and response.status_code < 400 and not path.startswith('/admin/jsi18n'):
                    forwarded = request.META.get('HTTP_X_FORWARDED_FOR')
                    ip = forwarded.split(',')[0].strip() if forwarded else request.META.get('REMOTE_ADDR')
                    AuditLog.objects.create(
                        user=user,
                        action='DOWNLOAD' if exported else 'CHANGE',
                        details=f"{request.method} {path}",
                        ip_address=ip,
                    )
        except Exception:
            pass                                     # logging must never break a page
        return response
