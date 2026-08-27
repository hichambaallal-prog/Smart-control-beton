import io
import json
import os
import urllib.parse
import qrcode
import streamlit as st
import streamlit.components.v1 as components

try:
    from offline_manager import insert_safe, is_online
    OFFLINE_SUPPORT = True
except ImportError:
    OFFLINE_SUPPORT = False


def _imprimer_qr_codes_js(eprouvettes, rec_id, beton_id, app_url_base):
    """
    Génère les QR codes et injecte du code JS avec fenêtre d'impression isolée.
    Correction complète de l'échappement des accolades Python/JS.
    """
    if not eprouvettes:
        st.warning("Aucune éprouvette à imprimer.")
        return

    items_html = []
    for ep in eprouvettes:
        num_ep = ep.get("num", "")
        echeance = ep.get("echeance", "")
        date_ec = ep.get("date_ec", "")

        qr_param_str = f"rec={urllib.parse.quote(str(rec_id))}&beton_id={urllib.parse.quote(str(beton_id))}&ep={urllib.parse.quote(str(num_ep))}"
        qr_url = f"{app_url_base}/?{qr_param_str}"

        qr_img = qrcode.make(qr_url)
        buf = io.BytesIO()
        qr_img.save(buf, format="PNG")
        import base64
        qr_b64 = base64.b64encode(buf.getvalue()).decode("utf-8")

        items_html.append(f"""
        <div class="qr-card">
            <div class="title">LPEE - CTR-CSB</div>
            <img src="data:image/png;base64,{qr_b64}" class="qr-img" />
            <div class="code-rec">{rec_id}</div>
            <div class="info"><b>Béton :</b> {beton_id} | <b>Ép :</b> {num_ep}</div>
            <div class="info"><b>Échéance :</b> {echeance}j ({date_ec})</div>
        </div>
        """)

    cards_joined = "".join(items_html)

    # Note l'utilisation des doubles accolades {{ }} pour le CSS/JS dans le f-string
    print_script = f"""
    <script>
    function printQRCodes() {{
        var printWindow = window.open('', '_blank', 'width=800,height=600');
        var htmlContent = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>Impression QR Codes</title>
            <style>
                @page {{ size: A4; margin: 10mm; }}
                body {{ font-family: Arial, sans-serif; margin: 0; padding: 0; background: #fff; }}
                .grid {{ display: flex; flex-wrap: wrap; gap: 10px; padding: 10px; }}
                .qr-card {{
                    width: 45%;
                    border: 2px dashed #333;
                    border-radius: 8px;
                    padding: 8px;
                    text-align: center;
                    box-sizing: border-box;
                    page-break-inside: avoid;
                }}
                .title {{ font-weight: bold; font-size: 14px; margin-bottom: 4px; color: #0066cc; }}
                .qr-img {{ width: 120px; height: 120px; margin: 4px auto; display: block; }}
                .code-rec {{ font-weight: bold; font-size: 13px; margin: 2px 0; }}
                .info {{ font-size: 11px; color: #333; }}
            </style>
        </head>
        <body>
            <div class="grid">
                {cards_joined}
            </div>
            <script>
                window.onload = function() {{
                    window.print();
                    setTimeout(function() {{ window.close(); }}, 500);
                }};
            <\\/script>
        </body>
        </html>
        `;
        printWindow.document.write(htmlContent);
        printWindow.document.close();
    }}
    printQRCodes();
    </script>
    """
    components.html(print_script, height=0)


def show(supabase, can_edit=False, user=None):
    st.title("🧪 Suivi Contrôle Béton - Saisie & Éprouvettes")
    st.caption("Module d'enregistrement des essais et suivi du mûrissement.")

    if not can_edit:
        st.warning("🔒 Mode Consultation seul (Modification restreinte).")

    st.info("Interface prête pour la saisie et la gestion des contrôles béton.")
