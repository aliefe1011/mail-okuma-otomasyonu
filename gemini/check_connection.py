"""Explicit live check: sends synthetic data only, never mailbox content."""
from bridge import analyze

if __name__ == '__main__':
    request = {'body': '4500123456 siparis 10 kalem 12 Kasim 2026 tarihinde teslim edilecek.',
               'received_at': '2026-11-04',
               'candidates': [{'row': 2, 'order': '4500123456', 'item': '10', 'material': 'ABC1234'}]}
    try:
        rows = analyze(request)
        if not rows or rows[0][:3] != (2, '2026-11-12', 'Teslim Edilecek'):
            raise ValueError('Synthetic reply was not correctly interpreted')
    except Exception as exc:
        print('Connection/interpretation check FAILED:', type(exc).__name__)
        raise SystemExit(1)
    print('Gemini connection and synthetic reply interpretation OK.')
