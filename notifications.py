"""
Sistema de Notificaciones UNIAGRARIA - Facatativá 2026
Maneja emails, notificaciones en tiempo real y alertas
"""

import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import logging
from threading import Thread

logger = logging.getLogger(__name__)

class NotificationService:
    """Servicio central de notificaciones"""
    
    def __init__(self, app=None):
        self.app = app
        if app:
            self.init_app(app)
    
    def init_app(self, app):
        self.smtp_server = app.config.get('MAIL_SERVER', 'smtp.gmail.com')
        self.smtp_port = app.config.get('MAIL_PORT', 587)
        self.smtp_username = app.config.get('MAIL_USERNAME')
        self.smtp_password = app.config.get('MAIL_PASSWORD')
        self.from_email = app.config.get('MAIL_DEFAULT_SENDER', 'noreply@uniagraria.edu.co')
    
    def send_email_async(self, app, msg):
        """Enviar email de forma asíncrona"""
        with app.app_context():
            try:
                server = smtplib.SMTP(self.smtp_server, self.smtp_port)
                server.starttls()
                if self.smtp_username and self.smtp_password:
                    server.login(self.smtp_username, self.smtp_password)
                server.send_message(msg)
                server.quit()
                logger.info(f"✅ Email enviado a {msg['To']}")
            except Exception as e:
                logger.error(f"❌ Error enviando email: {str(e)}")
    
    def send_email(self, subject, recipients, body, html_body=None, attachments=None):
        """Enviar email a múltiples destinatarios"""
        if not self.smtp_server:
            logger.warning("Servidor SMTP no configurado")
            return False
        
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"[UNIAGRARIA 2026] {subject}"
        msg['From'] = self.from_email
        msg['To'] = ', '.join(recipients) if isinstance(recipients, list) else recipients
        
        # Adjuntar texto plano
        msg.attach(MIMEText(body, 'plain'))
        
        # Adjuntar HTML si existe
        if html_body:
            msg.attach(MIMEText(html_body, 'html'))
        
        # Adjuntar archivos
        if attachments:
            for attachment in attachments:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment['content'])
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{attachment["filename"]}"'
                )
                msg.attach(part)
        
        # Enviar en segundo plano
        Thread(target=self.send_email_async, args=(self.app._get_current_object(), msg)).start()
        return True
    
    def notify_new_registration(self, registro, usuario):
        """Notificar nuevo registro de recolección"""
        subject = f"Nuevo Registro - {registro.realizador_nombre} {registro.realizador_apellidos or ''}"
        
        body = f"""
        Se ha creado un nuevo registro en el sistema:
        
        ID: {registro.id}
        Estudiante: {registro.realizador_nombre} {registro.realizador_apellidos or ''}
        Documento: {registro.matricula_documento}
        Municipio: {registro.municipio.nombre if registro.municipio else 'N/A'}
        Institución: {registro.institucion.nombre if registro.institucion else 'N/A'}
        Fecha: {registro.fecha_registro.strftime('%d/%m/%Y %H:%M')}
        Usuario: {usuario.nombre}
        
        Accede al sistema para ver más detalles.
        """
        
        # Notificar a administradores
        from app import Usuario
        admins = Usuario.query.filter_by(rol='admin').all()
        admin_emails = [a.email for a in admins if a.email]
        
        if admin_emails:
            self.send_email(subject, admin_emails, body)
    
    def notify_import_complete(self, archivo, usuario):
        """Notificar importación completada"""
        subject = f"Importación Completada - {archivo.nombre_archivo}"
        
        body = f"""
        La importación de datos ha sido completada exitosamente:
        
        Archivo: {archivo.nombre_archivo}
        Registros procesados: {archivo.registros_procesados}
        Fecha: {archivo.fecha_importacion.strftime('%d/%m/%Y %H:%M')}
        Usuario: {usuario.nombre}
        
        URL del archivo: {archivo.url}
        """
        
        self.send_email(subject, usuario.email, body)
        
        # También notificar a otros admins
        from app import Usuario
        admins = Usuario.query.filter(Usuario.rol == 'admin', Usuario.id != usuario.id).all()
        admin_emails = [a.email for a in admins if a.email]
        
        if admin_emails:
            self.send_email(subject, admin_emails, body)

# Instancia global
notification_service = NotificationService()