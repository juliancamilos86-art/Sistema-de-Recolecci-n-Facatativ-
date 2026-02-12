"""
UNIAGRARIA - SISTEMA DE RECOLECCIÓN FACATATIVÁ 2026
VERSIÓN MEGA PROFESIONAL - AZURE AD + CLOUDINARY + NEONTECH
Desarrollado para optimizar procesos de recolección de datos
"""

import os
import json
import io
import csv
import zipfile
import hashlib
import hmac
import base64
import secrets
from datetime import datetime, timedelta
from functools import wraps
from io import BytesIO, StringIO
from urllib.parse import quote, urlencode

# Flask y extensiones core
from flask import (
    Flask, render_template, request, redirect, url_for, 
    session, jsonify, flash, send_file, make_response, abort
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, current_user, login_required
from flask_migrate import Migrate
from flask_session import Session
from flask_cors import CORS
from flask_compress import Compress
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_caching import Cache
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.security import generate_password_hash, check_password_hash

# Base de datos
from sqlalchemy import create_engine, text, func, and_, or_, desc
from sqlalchemy.dialects.postgresql import JSON, UUID, ARRAY
from sqlalchemy.exc import SQLAlchemyError, IntegrityError
from sqlalchemy.orm import declarative_base, relationship

# Azure AD
import msal
import requests
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient

# Cloudinary
import cloudinary
import cloudinary.uploader
import cloudinary.api
from cloudinary.uploader import upload
from cloudinary.utils import cloudinary_url

# Procesamiento de datos - SIN PANDAS NI NUMPY
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
<<<<<<< HEAD
=======

>>>>>>> 7bd5c01971cd39864a1e3a08956d12dd862f1613
from PIL import Image

# PDF y Reportes
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from io import BytesIO

# Utilidades
import re
import uuid
from email_validator import validate_email, EmailNotValidError
from werkzeug.utils import secure_filename

# Configuración de variables de entorno
from dotenv import load_dotenv
load_dotenv()

# ============================================================================
# CONFIGURACIÓN INICIAL DE LA APLICACIÓN
# ===========================================================================

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)

# Configuración desde variables de entorno
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', secrets.token_hex(32))
app.config['SESSION_TYPE'] = os.environ.get('SESSION_TYPE', 'filesystem')
app.config['SESSION_PERMANENT'] = True
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=24)
app.config['SESSION_USE_SIGNER'] = True
app.config['SESSION_COOKIE_SECURE'] = False  # Cambiar a True en producción con HTTPS
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

# Configuración de base de datos
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    'pool_size': 10,
    'pool_recycle': 300,
    'pool_pre_ping': True,
    'pool_use_lifo': True,
    'max_overflow': 20,
    'connect_args': {'sslmode': 'require', 'connect_timeout': 10}
}

# Configuración de archivos
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'pdf', 'xlsx', 'xls', 'csv'}
app.config['ALLOWED_IMAGES'] = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'}

# Configuración Cloudinary
cloudinary.config(
    cloud_name=os.environ.get('CLOUDINARY_CLOUD_NAME'),
    api_key=os.environ.get('CLOUDINARY_API_KEY'),
    api_secret=os.environ.get('CLOUDINARY_API_SECRET'),
    secure=True
)

# Inicializar extensiones
db = SQLAlchemy(app)
migrate = Migrate(app, db)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Por favor inicia sesión para acceder'
login_manager.session_protection = 'strong'
Session(app)
CORS(app, supports_credentials=True)
Compress(app)
cache = Cache(app, config={'CACHE_TYPE': 'SimpleCache', 'CACHE_DEFAULT_TIMEOUT': 300})

# Rate Limiting
limiter = Limiter(
    app=app,
    key_func=get_remote_address,
    default_limits=["200 per day", "50 per hour"],
    storage_uri=os.environ.get('REDIS_URL', "memory://"),
    strategy="fixed-window"
)

# ============================================================================
# MODELOS DE BASE DE DATOS (NeonTech PostgreSQL)
# ============================================================================

class Usuario(UserMixin, db.Model):
    """Modelo de usuarios con soporte para Azure AD"""
    __tablename__ = 'usuarios'
    
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(255), unique=True, nullable=False, index=True)
    nombre = db.Column(db.String(255), nullable=False)
    apellido = db.Column(db.String(255))
    azure_id = db.Column(db.String(255), unique=True)
    rol = db.Column(db.String(50), default='usuario')
    activo = db.Column(db.Boolean, default=True)
    avatar_url = db.Column(db.String(500))
    ultimo_acceso = db.Column(db.DateTime)
    fecha_registro = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relaciones
    registros = db.relationship('RecoleccionDato', backref='usuario_registro', lazy=True)
    imagenes_subidas = db.relationship('FeriaImagen', backref='usuario_subida', lazy=True)
    importaciones = db.relationship('ArchivoImportado', backref='usuario_importo', lazy=True)
    
    def is_admin(self):
        return self.rol == 'admin'
    
    def to_dict(self):
        return {
            'id': self.id,
            'email': self.email,
            'nombre': self.nombre,
            'apellido': self.apellido,
            'rol': self.rol,
            'activo': self.activo,
            'avatar': self.avatar_url,
            'ultimo_acceso': self.ultimo_acceso.isoformat() if self.ultimo_acceso else None
        }

class Municipio(db.Model):
    """Modelo de municipios de Cundinamarca"""
    __tablename__ = 'municipios'
    
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), unique=True, nullable=False, index=True)
    departamento = db.Column(db.String(100), default='Cundinamarca')
    activo = db.Column(db.Boolean, default=True)
    
    # Relaciones
    instituciones = db.relationship('Institucion', backref='municipio', lazy=True)
    registros = db.relationship('RecoleccionDato', backref='municipio', lazy=True)
    
    def to_dict(self):
        return {
            'id': self.id,
            'nombre': self.nombre,
            'departamento': self.departamento
        }

class Institucion(db.Model):
    """Modelo de instituciones educativas"""
    __tablename__ = 'instituciones'
    
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(255), nullable=False, index=True)
    municipio_id = db.Column(db.Integer, db.ForeignKey('municipios.id'), nullable=False)
    direccion = db.Column(db.String(255))
    telefono = db.Column(db.String(50))
    contacto = db.Column(db.String(255))
    activo = db.Column(db.Boolean, default=True)
    
    # Relaciones
    registros = db.relationship('RecoleccionDato', backref='institucion', lazy=True)
    
    def to_dict(self):
        return {
            'id': self.id,
            'nombre': self.nombre,
            'municipio_id': self.municipio_id,
            'municipio': self.municipio.nombre if self.municipio else None,
            'direccion': self.direccion,
            'telefono': self.telefono,
            'contacto': self.contacto
        }

class RecoleccionDato(db.Model):
    """Modelo principal para la recolección de datos 2026-1"""
    __tablename__ = 'recoleccion_datos'
    
    id = db.Column(db.Integer, primary_key=True)
    ficha_toma_registro = db.Column(db.String(50))
    asesor = db.Column(db.String(255))
    municipio_id = db.Column(db.Integer, db.ForeignKey('municipios.id'))
    institucion_id = db.Column(db.Integer, db.ForeignKey('instituciones.id'))
    realizador_nombre = db.Column(db.String(255))
    realizador_apellidos = db.Column(db.String(255))
    matricula_documento = db.Column(db.String(100))
    telefono = db.Column(db.String(50))
    correo = db.Column(db.String(255))
    grado = db.Column(db.String(50))
    instalacion_educativa = db.Column(db.String(255))
    pista = db.Column(db.String(100))
    programa_interes = db.Column(db.String(255))
    jornada_interes = db.Column(db.String(100))
    pendido_interes = db.Column(db.Text)
    asesoria_migradora = db.Column(db.Text)
    estado = db.Column(db.String(50), default='pendiente')
    ano_periodo = db.Column(db.String(50), default='2026-1')
    observacion = db.Column(db.Text)
    usuario_registro_id = db.Column(db.Integer, db.ForeignKey('usuarios.id'))
    fecha_registro = db.Column(db.DateTime, default=datetime.utcnow)
    fecha_actualizacion = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def to_dict(self):
        return {
            'id': self.id,
            'ficha': self.ficha_toma_registro,
            'asesor': self.asesor,
            'municipio': self.municipio.nombre if self.municipio else None,
            'municipio_id': self.municipio_id,
            'institucion': self.institucion.nombre if self.institucion else None,
            'institucion_id': self.institucion_id,
            'nombre': self.realizador_nombre,
            'apellidos': self.realizador_apellidos,
            'matricula': self.matricula_documento,
            'telefono': self.telefono,
            'correo': self.correo,
            'grado': self.grado,
            'pista': self.pista,
            'programa': self.programa_interes,
            'jornada': self.jornada_interes,
            'pendido': self.pendido_interes,
            'asesoria': self.asesoria_migradora,
            'estado': self.estado,
            'periodo': self.ano_periodo,
            'observacion': self.observacion,
            'fecha': self.fecha_registro.isoformat() if self.fecha_registro else None,
            'usuario': self.usuario_registro.nombre if self.usuario_registro else None
        }

class Feria(db.Model):
    """Modelo de ferias educativas"""
    __tablename__ = 'ferias'
    
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(255), nullable=False)
    fecha_inicio = db.Column(db.Date)
    fecha_fin = db.Column(db.Date)
    ubicacion = db.Column(db.String(255))
    municipio_id = db.Column(db.Integer, db.ForeignKey('municipios.id'))
    descripcion = db.Column(db.Text)
    activa = db.Column(db.Boolean, default=True)
    
    # Relaciones
    imagenes = db.relationship('FeriaImagen', backref='feria', lazy=True)
    municipio_rel = db.relationship('Municipio')
    
    def to_dict(self):
        return {
            'id': self.id,
            'nombre': self.nombre,
            'fecha_inicio': self.fecha_inicio.isoformat() if self.fecha_inicio else None,
            'fecha_fin': self.fecha_fin.isoformat() if self.fecha_fin else None,
            'ubicacion': self.ubicacion,
            'municipio': self.municipio_rel.nombre if self.municipio_rel else None,
            'descripcion': self.descripcion,
            'activa': self.activa,
            'total_imagenes': len(self.imagenes) if self.imagenes else 0
        }

class FeriaImagen(db.Model):
    """Modelo de imágenes de ferias (Cloudinary)"""
    __tablename__ = 'ferias_imagenes'
    
    id = db.Column(db.Integer, primary_key=True)
    feria_id = db.Column(db.Integer, db.ForeignKey('ferias.id'))
    public_id = db.Column(db.String(255))
    url = db.Column(db.String(500))
    descripcion = db.Column(db.Text)
    usuario_subida_id = db.Column(db.Integer, db.ForeignKey('usuarios.id'))
    fecha_subida = db.Column(db.DateTime, default=datetime.utcnow)
    
    def to_dict(self):
        return {
            'id': self.id,
            'feria_id': self.feria_id,
            'public_id': self.public_id,
            'url': self.url,
            'descripcion': self.descripcion,
            'usuario': self.usuario_subida.nombre if self.usuario_subida else None,
            'fecha': self.fecha_subida.isoformat() if self.fecha_subida else None
        }

class ArchivoImportado(db.Model):
    """Modelo para tracking de archivos importados"""
    __tablename__ = 'archivos_importados'
    
    id = db.Column(db.Integer, primary_key=True)
    nombre_archivo = db.Column(db.String(255))
    tipo = db.Column(db.String(50))
    url = db.Column(db.String(500))
    usuario_importo_id = db.Column(db.Integer, db.ForeignKey('usuarios.id'))
    fecha_importacion = db.Column(db.DateTime, default=datetime.utcnow)
    registros_procesados = db.Column(db.Integer, default=0)
    estado = db.Column(db.String(50), default='completado')
    datos_metadata = db.Column(db.JSON, default={})

# ============================================================================
# FUNCIONES AUXILIARES Y UTILITARIAS
# ============================================================================

def allowed_file(filename, file_types=None):
    """Verifica si el archivo tiene extensión permitida"""
    if '.' not in filename:
        return False
    extension = filename.rsplit('.', 1)[1].lower()
    if file_types == 'image':
        return extension in app.config['ALLOWED_IMAGES']
    return extension in app.config['ALLOWED_EXTENSIONS']

def validate_email_format(email):
    """Valida formato de email"""
    try:
        valid = validate_email(email)
        return valid.email
    except EmailNotValidError:
        return None

def generar_excel_recoleccion(filtros=None):
    """Genera archivo Excel con datos de recolección - SIN PANDAS"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    
    query = RecoleccionDato.query
    if filtros:
        if filtros.get('municipio_id'):
            query = query.filter_by(municipio_id=filtros['municipio_id'])
        if filtros.get('ano_periodo'):
            query = query.filter_by(ano_periodo=filtros['ano_periodo'])
        if filtros.get('estado'):
            query = query.filter_by(estado=filtros['estado'])
        if filtros.get('fecha_inicio') and filtros.get('fecha_fin'):
            query = query.filter(
                RecoleccionDato.fecha_registro.between(
                    filtros['fecha_inicio'], filtros['fecha_fin']
                )
            )
    
    datos = query.order_by(RecoleccionDato.fecha_registro.desc()).all()
    
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Recolección 2026-1"
    
    # Encabezados
    headers = ['ID', 'Ficha', 'Asesor', 'Municipio', 'Institución', 'Nombre', 
               'Apellidos', 'Documento', 'Teléfono', 'Email', 'Grado', 'Programa', 
               'Jornada', 'Estado', 'Periodo', 'Observación', 'Fecha Registro']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    
    # Datos
    for row, d in enumerate(datos, 2):
        ws.cell(row=row, column=1, value=d.id)
        ws.cell(row=row, column=2, value=d.ficha_toma_registro)
        ws.cell(row=row, column=3, value=d.asesor)
        ws.cell(row=row, column=4, value=d.municipio.nombre if d.municipio else '')
        ws.cell(row=row, column=5, value=d.institucion.nombre if d.institucion else '')
        ws.cell(row=row, column=6, value=d.realizador_nombre)
        ws.cell(row=row, column=7, value=d.realizador_apellidos)
        ws.cell(row=row, column=8, value=d.matricula_documento)
        ws.cell(row=row, column=9, value=d.telefono)
        ws.cell(row=row, column=10, value=d.correo)
        ws.cell(row=row, column=11, value=d.grado)
        ws.cell(row=row, column=12, value=d.programa_interes)
        ws.cell(row=row, column=13, value=d.jornada_interes)
        ws.cell(row=row, column=14, value=d.estado)
        ws.cell(row=row, column=15, value=d.ano_periodo)
        ws.cell(row=row, column=16, value=d.observacion)
        ws.cell(row=row, column=17, value=d.fecha_registro.strftime('%Y-%m-%d %H:%M') if d.fecha_registro else '')
    
    # Ajustar ancho de columnas
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(output)
    output.seek(0)
    return output

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(Usuario, int(user_id))

def admin_required(f):
    """Decorador para rutas que requieren ser administrador"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_admin():
            flash('Acceso denegado. Se requieren permisos de administrador.', 'error')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

# ============================================================================
# RUTAS DE AUTENTICACIÓN AZURE AD
# ============================================================================

@app.route('/login')
def login():
    """Redirige a Microsoft Azure AD para autenticación"""
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    
    session['oauth_state'] = secrets.token_urlsafe(32)
    
    params = {
        'client_id': os.environ.get('AZURE_CLIENT_ID'),
        'response_type': 'code id_token',
        'redirect_uri': os.environ.get('AZURE_REDIRECT_URI'),
        'response_mode': 'form_post',
        'scope': 'openid profile email User.Read',
        'state': session['oauth_state'],
        'nonce': secrets.token_urlsafe(32)
    }
    
    auth_url = f"https://login.microsoftonline.com/{os.environ.get('AZURE_TENANT_ID')}/oauth2/v2.0/authorize?{urlencode(params)}"
    return redirect(auth_url)

@app.route('/callback', methods=['POST'])
def callback():
    """Callback de Azure AD"""
    try:
        if request.form.get('state') != session.get('oauth_state'):
            flash('Error de autenticación: estado inválido', 'error')
            return redirect(url_for('login'))
        
        tenant_id = os.environ.get('AZURE_TENANT_ID')
        client_id = os.environ.get('AZURE_CLIENT_ID')
        client_secret = os.environ.get('AZURE_CLIENT_SECRET')
        
        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        
        token_data = {
            'client_id': client_id,
            'client_secret': client_secret,
            'code': request.form.get('code'),
            'redirect_uri': os.environ.get('AZURE_REDIRECT_URI'),
            'grant_type': 'authorization_code'
        }
        
        token_response = requests.post(token_url, data=token_data)
        token_json = token_response.json()
        
        access_token = token_json.get('access_token')
        headers = {'Authorization': f'Bearer {access_token}'}
        graph_response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
        graph_data = graph_response.json()
        
        email = graph_data.get('mail') or graph_data.get('userPrincipalName')
        nombre = graph_data.get('displayName', 'Usuario')
        azure_id = graph_data.get('id')
        
        usuario = Usuario.query.filter_by(email=email).first()
        
        if not usuario:
            total_usuarios = Usuario.query.count()
            rol = 'admin' if total_usuarios < 5 else 'usuario'
            
            usuario = Usuario(
                email=email,
                nombre=nombre,
                azure_id=azure_id,
                rol=rol,
                activo=True
            )
            db.session.add(usuario)
            db.session.commit()
            
            flash('¡Bienvenido a Uniagraria Sistema 2026!', 'success')
        else:
            usuario.ultimo_acceso = datetime.utcnow()
            db.session.commit()
        
        login_user(usuario, remember=True)
        session.permanent = True
        
        return redirect(url_for('dashboard'))
        
    except Exception as e:
        app.logger.error(f"Error en callback Azure AD: {str(e)}")
        flash('Error en la autenticación con Azure AD', 'error')
        return redirect(url_for('login'))

@app.route('/logout')
@login_required
def logout():
    """Cerrar sesión"""
    logout_user()
    session.clear()
    
    logout_url = f"https://login.microsoftonline.com/{os.environ.get('AZURE_TENANT_ID')}/oauth2/v2.0/logout"
    params = {
        'post_logout_redirect_uri': url_for('login', _external=True)
    }
    return redirect(f"{logout_url}?{urlencode(params)}")

# ============================================================================
# RUTAS PRINCIPALES
# ============================================================================

@app.route('/')
def index():
    """Redirigir a dashboard si está autenticado, sino a login"""
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
@cache.cached(timeout=60)
def dashboard():
    """Dashboard principal con estadísticas - SIN PANDAS NI PLOTLY"""
    try:
        total_registros = RecoleccionDato.query.filter_by(ano_periodo='2026-1').count()
        total_municipios = db.session.query(func.count(db.distinct(RecoleccionDato.municipio_id))).filter_by(ano_periodo='2026-1').scalar() or 0
        total_instituciones = db.session.query(func.count(db.distinct(RecoleccionDato.institucion_id))).filter_by(ano_periodo='2026-1').scalar() or 0
        completados = RecoleccionDato.query.filter_by(ano_periodo='2026-1', estado='completado').count()
        pendientes = RecoleccionDato.query.filter_by(ano_periodo='2026-1', estado='pendiente').count()
        
        registros_recientes = RecoleccionDato.query.order_by(
            RecoleccionDato.fecha_registro.desc()
        ).limit(10).all()
        
        return render_template('dashboard.html',
                             total_registros=total_registros,
                             total_municipios=total_municipios,
                             total_instituciones=total_instituciones,
                             completados=completados,
                             pendientes=pendientes,
                             registros_recientes=registros_recientes,
                             usuario=current_user,
                             now=datetime.now)
    except Exception as e:
        app.logger.error(f"Error en dashboard: {str(e)}")
        flash('Error al cargar el dashboard', 'error')
        return render_template('dashboard.html', usuario=current_user, now=datetime.now)

# ============================================================================
# RUTAS DE RECOLECCIÓN DE DATOS
# ============================================================================

@app.route('/recoleccion')
@login_required
def recoleccion():
    """Vista de recolección de datos"""
    municipios = Municipio.query.filter_by(activo=True).all()
    instituciones = Institucion.query.filter_by(activo=True).all()
    
    municipio_id = request.args.get('municipio', type=int)
    periodo = request.args.get('periodo', '2026-1')
    
    query = RecoleccionDato.query.filter_by(ano_periodo=periodo)
    
    if municipio_id:
        query = query.filter_by(municipio_id=municipio_id)
    
    registros = query.order_by(RecoleccionDato.fecha_registro.desc()).limit(100).all()
    
    return render_template('recoleccion.html',
                         registros=registros,
                         municipios=municipios,
                         instituciones=instituciones,
                         periodo_actual=periodo)

@app.route('/api/recoleccion', methods=['POST'])
@login_required
def api_crear_recoleccion():
    """API para crear nuevo registro de recolección"""
    try:
        data = request.get_json()
        
        if data.get('correo'):
            email_validado = validate_email_format(data['correo'])
            if not email_validado:
                return jsonify({'error': 'Email inválido'}), 400
            data['correo'] = email_validado
        
        nuevo_registro = RecoleccionDato(
            ficha_toma_registro=data.get('ficha_toma_registro'),
            asesor=data.get('asesor'),
            municipio_id=data.get('municipio_id'),
            institucion_id=data.get('institucion_id'),
            realizador_nombre=data.get('realizador_nombre'),
            realizador_apellidos=data.get('realizador_apellidos'),
            matricula_documento=data.get('matricula_documento'),
            telefono=data.get('telefono'),
            correo=data.get('correo'),
            grado=data.get('grado'),
            instalacion_educativa=data.get('instalacion_educativa'),
            pista=data.get('pista'),
            programa_interes=data.get('programa_interes'),
            jornada_interes=data.get('jornada_interes'),
            pendido_interes=data.get('pendido_interes'),
            asesoria_migradora=data.get('asesoria_migradora'),
            estado=data.get('estado', 'pendiente'),
            ano_periodo=data.get('ano_periodo', '2026-1'),
            observacion=data.get('observacion'),
            usuario_registro_id=current_user.id
        )
        
        db.session.add(nuevo_registro)
        db.session.commit()
        
        return jsonify({
            'message': 'Registro creado exitosamente',
            'data': nuevo_registro.to_dict()
        }), 201
        
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Error al crear registro: {str(e)}")
        return jsonify({'error': 'Error al crear el registro'}), 500

@app.route('/api/recoleccion/<int:id>', methods=['PUT'])
@login_required
def api_actualizar_recoleccion(id):
    """API para actualizar registro de recolección"""
    try:
        registro = RecoleccionDato.query.get_or_404(id)
        data = request.get_json()
        
        if data.get('correo'):
            email_validado = validate_email_format(data['correo'])
            if not email_validado:
                return jsonify({'error': 'Email inválido'}), 400
            data['correo'] = email_validado
        
        for key, value in data.items():
            if hasattr(registro, key):
                setattr(registro, key, value)
        
        registro.fecha_actualizacion = datetime.utcnow()
        db.session.commit()
        
        return jsonify({
            'message': 'Registro actualizado exitosamente',
            'data': registro.to_dict()
        })
        
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Error al actualizar registro: {str(e)}")
        return jsonify({'error': 'Error al actualizar el registro'}), 500

# ============================================================================
# RUTAS DE IMPORTACIÓN DE DATOS - SIN PANDAS
# ============================================================================

@app.route('/importacion')
@login_required
@admin_required
def importacion():
    """Página de importación de datos"""
    importaciones = ArchivoImportado.query.order_by(
        ArchivoImportado.fecha_importacion.desc()
    ).limit(20).all()
    
    return render_template('importacion.html', 
                         importaciones=importaciones)

@app.route('/api/importacion/excel', methods=['POST'])
@login_required
@admin_required
@limiter.limit("10 per minute")
def api_importar_excel():
    """Importar archivo Excel con datos de recolección - SIN PANDAS"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No se envió ningún archivo'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'Nombre de archivo vacío'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Tipo de archivo no permitido. Use Excel (.xlsx, .xls)'}), 400
        
        filename = secure_filename(file.filename)
        
        # Usar openpyxl directamente en lugar de pandas
        from openpyxl import load_workbook
        
        try:
            wb = load_workbook(file)
            ws = wb.active
            headers = [cell.value for cell in ws[1]]
            data = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any(cell is not None for cell in row):
                    data.append(dict(zip(headers, row)))
        except Exception as e:
            return jsonify({'error': f'Error al leer el archivo Excel: {str(e)}'}), 400
        
        if not data:
            return jsonify({'error': 'El archivo está vacío'}), 400
        
        registros_procesados = 0
        errores = []
        
        for idx, row in enumerate(data, 2):
            try:
                municipio_nombre = str(row.get('Municipio', row.get('municipio', ''))).strip()
                if municipio_nombre:
                    municipio = Municipio.query.filter(
                        func.lower(Municipio.nombre) == func.lower(municipio_nombre)
                    ).first()
                    
                    if not municipio:
                        municipio = Municipio(nombre=municipio_nombre)
                        db.session.add(municipio)
                        db.session.flush()
                else:
                    municipio = None
                
                institucion_nombre = str(row.get('Instalacion_Educativa', row.get('instalacion_educativa', ''))).strip()
                if institucion_nombre and municipio:
                    institucion = Institucion.query.filter(
                        func.lower(Institucion.nombre) == func.lower(institucion_nombre),
                        Institucion.municipio_id == municipio.id
                    ).first()
                    
                    if not institucion:
                        institucion = Institucion(
                            nombre=institucion_nombre,
                            municipio_id=municipio.id
                        )
                        db.session.add(institucion)
                        db.session.flush()
                else:
                    institucion = None
                
                registro = RecoleccionDato(
                    ficha_toma_registro=str(row.get('Ficha_toma_registro', row.get('ficha', '')))[:50],
                    asesor=str(row.get('Asesor', row.get('asesor', '')))[:255],
                    municipio_id=municipio.id if municipio else None,
                    institucion_id=institucion.id if institucion else None,
                    realizador_nombre=str(row.get('Realizador_Nombre', row.get('nombre', '')))[:255],
                    realizador_apellidos=str(row.get('Realizador_Apellidos', row.get('apellidos', '')))[:255],
                    matricula_documento=str(row.get('Matricula_Documento', row.get('matricula', '')))[:100],
                    telefono=str(row.get('Telefono', row.get('telefono', '')))[:50],
                    correo=str(row.get('Correo', row.get('correo', '')))[:255],
                    grado=str(row.get('Grado', row.get('grado', '')))[:50],
                    instalacion_educativa=str(row.get('Instalacion_Educativa', row.get('instalacion_educativa', '')))[:255],
                    pista=str(row.get('Pista', row.get('pista', '')))[:100],
                    programa_interes=str(row.get('Programa_Interes', row.get('programa_interes', '')))[:255],
                    jornada_interes=str(row.get('Jornada_Interes', row.get('jornada_interes', '')))[:100],
                    pendido_interes=str(row.get('Pendido_Interes', row.get('pendido_interes', ''))),
                    asesoria_migradora=str(row.get('Asesoria_Migradora', row.get('asesoria_migradora', ''))),
                    estado=str(row.get('Estado', row.get('estado', 'pendiente'))),
                    ano_periodo=str(row.get('Ano_periodo', row.get('ano_periodo', '2026-1'))),
                    observacion=str(row.get('Observacion', row.get('observacion', ''))),
                    usuario_registro_id=current_user.id
                )
                
                db.session.add(registro)
                registros_procesados += 1
                
                if registros_procesados % 50 == 0:
                    db.session.commit()
                
            except Exception as e:
                db.session.rollback()
                errores.append({
                    'fila': idx,
                    'error': str(e)[:200]
                })
        
        try:
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            return jsonify({'error': f'Error al guardar registros: {str(e)}'}), 500
        
        try:
            file.seek(0)
            result = cloudinary.uploader.upload(
                file,
                folder='uniagraria/importaciones',
                public_id=f"importacion_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
                resource_type='raw'
            )
            
            archivo_importado = ArchivoImportado(
                nombre_archivo=filename,
                tipo='excel',
                url=result['secure_url'],
                usuario_importo_id=current_user.id,
                registros_procesados=registros_procesados,
                estado='completado',
                datos_metadata={
                    'total_filas': len(data),
                    'procesados': registros_procesados,
                    'errores': len(errores),
                    'fecha': datetime.now().isoformat()
                }
            )
            db.session.add(archivo_importado)
            db.session.commit()
            
        except Exception as e:
            app.logger.error(f"Error al subir a Cloudinary: {str(e)}")
        
        return jsonify({
            'message': f'Importación completada: {registros_procesados} registros procesados',
            'registros_procesados': registros_procesados,
            'errores': len(errores),
            'detalles_errores': errores[:10]
        })
        
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Error general en importación: {str(e)}")
        return jsonify({'error': 'Error interno en el servidor'}), 500

# ============================================================================
# RUTAS DE REPORTES - SIN PANDAS
# ============================================================================

@app.route('/reportes')
@login_required
def reportes():
    """Página de reportes avanzados"""
    municipios = Municipio.query.filter_by(activo=True).all()
    return render_template('reportes.html', municipios=municipios)

@app.route('/api/reportes/general')
@login_required
def api_reporte_general():
    """API para reporte general en JSON - SIN PANDAS"""
    try:
        periodo = request.args.get('periodo', '2026-1')
        municipio_id = request.args.get('municipio_id', type=int)
        
        query = RecoleccionDato.query.filter_by(ano_periodo=periodo)
        
        if municipio_id:
            query = query.filter_by(municipio_id=municipio_id)
        
        registros = query.all()
        
        stats = {
            'total': len(registros),
            'completados': sum(1 for r in registros if r.estado == 'completado'),
            'pendientes': sum(1 for r in registros if r.estado == 'pendiente'),
            'por_grado': {},
            'por_programa': {},
            'por_municipio': {}
        }
        
        for r in registros:
            if r.grado:
                stats['por_grado'][r.grado] = stats['por_grado'].get(r.grado, 0) + 1
            if r.programa_interes:
                stats['por_programa'][r.programa_interes] = stats['por_programa'].get(r.programa_interes, 0) + 1
            if r.municipio:
                nombre_municipio = r.municipio.nombre
                stats['por_municipio'][nombre_municipio] = stats['por_municipio'].get(nombre_municipio, 0) + 1
        
        return jsonify({
            'periodo': periodo,
            'estadisticas': stats,
            'registros': [r.to_dict() for r in registros[:100]]
        })
        
    except Exception as e:
        app.logger.error(f"Error en reporte general: {str(e)}")
        return jsonify({'error': 'Error al generar reporte'}), 500

@app.route('/api/reportes/exportar/excel')
@login_required
def api_exportar_excel():
    """Exportar datos a Excel"""
    try:
        filtros = {
            'municipio_id': request.args.get('municipio_id', type=int),
            'ano_periodo': request.args.get('periodo', '2026-1'),
            'estado': request.args.get('estado'),
            'fecha_inicio': request.args.get('fecha_inicio'),
            'fecha_fin': request.args.get('fecha_fin')
        }
        
        output = generar_excel_recoleccion(filtros)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'reporte_uniagraria_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
        
    except Exception as e:
        app.logger.error(f"Error al exportar Excel: {str(e)}")
        flash('Error al generar el archivo Excel', 'error')
        return redirect(url_for('reportes'))

@app.route('/api/reportes/exportar/pdf')
@login_required
def api_exportar_pdf():
    """Exportar reporte a PDF profesional"""
    try:
        periodo = request.args.get('periodo', '2026-1')
        
        registros = RecoleccionDato.query.filter_by(ano_periodo=periodo).all()
        total_registros = len(registros)
        
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        elements = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#2E7D32'),
            spaceAfter=30,
            alignment=1
        )
        
        elements.append(Paragraph(f'UNIAGRARIA - RECOLECCIÓN FACATATIVÁ {periodo}', title_style))
        elements.append(Spacer(1, 20))
        
        info_style = ParagraphStyle(
            'InfoStyle',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=10
        )
        
        elements.append(Paragraph(f'Fecha de generación: {datetime.now().strftime("%d/%m/%Y %H:%M")}', info_style))
        elements.append(Paragraph(f'Total de registros: {total_registros}', info_style))
        elements.append(Paragraph(f'Generado por: {current_user.nombre}', info_style))
        elements.append(Spacer(1, 30))
        
        data = [['ID', 'Nombre', 'Municipio', 'Institución', 'Grado', 'Estado']]
        
        for r in registros[:50]:
            data.append([
                str(r.id),
                f"{r.realizador_nombre} {r.realizador_apellidos or ''}",
                r.municipio.nombre if r.municipio else 'N/A',
                r.institucion.nombre[:30] if r.institucion else 'N/A',
                r.grado or 'N/A',
                r.estado
            ])
        
        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E7D32')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elements.append(table)
        
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'reporte_uniagraria_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
        )
        
    except Exception as e:
        app.logger.error(f"Error al exportar PDF: {str(e)}")
        flash('Error al generar el PDF', 'error')
        return redirect(url_for('reportes'))

# ============================================================================
# RUTAS DE FERIAS - IMÁGENES Y CLOUDINARY
# ============================================================================

@app.route('/ferias')
@login_required
def ferias():
    """Página de gestión de ferias"""
    ferias_list = Feria.query.filter_by(activa=True).all()
    municipios = Municipio.query.filter_by(activo=True).all()
    return render_template('ferias.html', 
                         ferias=ferias_list, 
                         municipios=municipios)

@app.route('/api/ferias', methods=['POST'])
@login_required
def api_crear_feria():
    """Crear nueva feria"""
    try:
        data = request.get_json()
        
        nueva_feria = Feria(
            nombre=data.get('nombre'),
            fecha_inicio=datetime.strptime(data.get('fecha_inicio'), '%Y-%m-%d') if data.get('fecha_inicio') else None,
            fecha_fin=datetime.strptime(data.get('fecha_fin'), '%Y-%m-%d') if data.get('fecha_fin') else None,
            ubicacion=data.get('ubicacion'),
            municipio_id=data.get('municipio_id'),
            descripcion=data.get('descripcion'),
            activa=True
        )
        
        db.session.add(nueva_feria)
        db.session.commit()
        
        return jsonify({
            'message': 'Feria creada exitosamente',
            'data': nueva_feria.to_dict()
        }), 201
        
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Error al crear feria: {str(e)}")
        return jsonify({'error': 'Error al crear la feria'}), 500

@app.route('/api/ferias/<int:feria_id>/imagenes', methods=['POST'])
@login_required
def api_subir_imagen_feria(feria_id):
    """Subir imágenes de feria a Cloudinary"""
    try:
        feria = Feria.query.get_or_404(feria_id)
        
        if 'images' not in request.files:
            return jsonify({'error': 'No se enviaron imágenes'}), 400
        
        files = request.files.getlist('images')
        imagenes_subidas = []
        errores = []
        
        for file in files:
            if file and allowed_file(file.filename, 'image'):
                try:
                    result = cloudinary.uploader.upload(
                        file,
                        folder=f'uniagraria/ferias/{feria_id}',
                        public_id=f"{feria_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{secrets.token_hex(4)}",
                        transformation=[
                            {'width': 1200, 'height': 800, 'crop': 'limit'},
                            {'quality': 'auto'}
                        ]
                    )
                    
                    imagen = FeriaImagen(
                        feria_id=feria_id,
                        public_id=result['public_id'],
                        url=result['secure_url'],
                        descripcion=request.form.get('descripcion', ''),
                        usuario_subida_id=current_user.id
                    )
                    
                    db.session.add(imagen)
                    imagenes_subidas.append({
                        'url': result['secure_url'],
                        'public_id': result['public_id']
                    })
                    
                except Exception as e:
                    errores.append({
                        'filename': file.filename,
                        'error': str(e)
                    })
        
        if imagenes_subidas:
            db.session.commit()
        
        return jsonify({
            'message': f'{len(imagenes_subidas)} imágenes subidas exitosamente',
            'imagenes': imagenes_subidas,
            'errores': errores
        })
        
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Error al subir imágenes: {str(e)}")
        return jsonify({'error': 'Error al subir las imágenes'}), 500

@app.route('/api/ferias/<int:feria_id>/imagenes')
@login_required
def api_obtener_imagenes_feria(feria_id):
    """Obtener imágenes de una feria"""
    try:
        imagenes = FeriaImagen.query.filter_by(feria_id=feria_id).order_by(
            FeriaImagen.fecha_subida.desc()
        ).all()
        
        return jsonify({
            'imagenes': [i.to_dict() for i in imagenes]
        })
        
    except Exception as e:
        app.logger.error(f"Error al obtener imágenes: {str(e)}")
        return jsonify({'error': 'Error al obtener imágenes'}), 500

# ============================================================================
# RUTAS DE ARCHIVOS Y COMPRESIÓN ZIP
# ============================================================================

@app.route('/archivos')
@login_required
def archivos():
    """Página de gestión de archivos"""
    importaciones = ArchivoImportado.query.order_by(
        ArchivoImportado.fecha_importacion.desc()
    ).limit(30).all()
    
    return render_template('archivos.html', 
                         importaciones=importaciones)

@app.route('/api/archivos/comprimir-todo')
@login_required
@admin_required
def api_comprimir_todo():
    """Comprimir todos los datos e imágenes en ZIP"""
    try:
        zip_buffer = BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            
            registros = RecoleccionDato.query.filter_by(ano_periodo='2026-1').all()
            datos_recoleccion = [r.to_dict() for r in registros]
            zip_file.writestr('recoleccion_datos_2026-1.json', 
                            json.dumps(datos_recoleccion, indent=2, default=str))
            
            ferias_list = Feria.query.all()
            datos_ferias = [f.to_dict() for f in ferias_list]
            zip_file.writestr('ferias.json', 
                            json.dumps(datos_ferias, indent=2, default=str))
            
            imagenes = FeriaImagen.query.all()
            datos_imagenes = [i.to_dict() for i in imagenes]
            zip_file.writestr('ferias_imagenes.json', 
                            json.dumps(datos_imagenes, indent=2, default=str))
            
            excel_buffer = generar_excel_recoleccion()
            zip_file.writestr('recoleccion_datos_2026-1.xlsx', excel_buffer.getvalue())
            
            pdf_buffer = BytesIO()
            doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()
            elements.append(Paragraph('UNIAGRARIA - REPORTE COMPLETO 2026-1', 
                                    styles['Title']))
            doc.build(elements)
            zip_file.writestr('reporte_completo_2026-1.pdf', pdf_buffer.getvalue())
            
            readme_content = f"""UNIAGRARIA - SISTEMA DE RECOLECCIÓN FACATATIVÁ 2026
================================================
Fecha de exportación: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
Usuario: {current_user.nombre} ({current_user.email})
Total registros: {len(registros)}
Total ferias: {len(ferias_list)}
Total imágenes: {len(imagenes)}

Este archivo contiene:
- recoleccion_datos_2026-1.json: Datos completos en formato JSON
- recoleccion_datos_2026-1.xlsx: Datos en formato Excel
- ferias.json: Información de ferias
- ferias_imagenes.json: URLs de imágenes en Cloudinary
- reporte_completo_2026-1.pdf: Reporte en PDF

Soporte: soporte@uniagraria.edu.co
"""
            zip_file.writestr('README.txt', readme_content)
        
        zip_buffer.seek(0)
        
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'uniagraria_backup_completo_{datetime.now().strftime("%Y%m%d_%H%M%S")}.zip'
        )
        
    except Exception as e:
        app.logger.error(f"Error al comprimir archivos: {str(e)}")
        flash('Error al generar el archivo ZIP', 'error')
        return redirect(url_for('archivos'))

# ============================================================================
# RUTAS DE ADMINISTRACIÓN
# ============================================================================

@app.route('/admin/usuarios')
@login_required
@admin_required
def admin_usuarios():
    """Panel de administración de usuarios"""
    usuarios = Usuario.query.order_by(Usuario.fecha_registro.desc()).all()
    return render_template('admin_usuarios.html', usuarios=usuarios)

@app.route('/api/admin/usuarios/<int:user_id>/rol', methods=['PUT'])
@login_required
@admin_required
def api_cambiar_rol(user_id):
    """Cambiar rol de usuario"""
    try:
        usuario = Usuario.query.get_or_404(user_id)
        data = request.get_json()
        
        nuevo_rol = data.get('rol')
        if nuevo_rol in ['admin', 'usuario']:
            usuario.rol = nuevo_rol
            db.session.commit()
            return jsonify({'message': 'Rol actualizado exitosamente'})
        
        return jsonify({'error': 'Rol inválido'}), 400
        
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Error al cambiar rol: {str(e)}")
        return jsonify({'error': 'Error al cambiar rol'}), 500

# ============================================================================
# MANEJO DE ERRORES
# ============================================================================

@app.errorhandler(404)
def not_found_error(error):
    return render_template('error.html', 
                         error_code=404, 
                         error_message='Página no encontrada'), 404

@app.errorhandler(500)
def internal_error(error):
    db.session.rollback()
    return render_template('error.html', 
                         error_code=500, 
                         error_message='Error interno del servidor'), 500

@app.errorhandler(429)
def ratelimit_error(error):
    return render_template('error.html',
                         error_code=429,
                         error_message='Demasiadas solicitudes. Por favor espere.'), 429

# ============================================================================
# INICIALIZACIÓN DE LA APLICACIÓN
# ============================================================================

def init_database():
    """Inicializar base de datos y datos por defecto"""
    with app.app_context():
        try:
            db.create_all()
            
            municipios_iniciales = [
                'Facatativá', 'Bogotá', 'Madrid', 'Mosquera', 
                'Funza', 'El Rosal', 'Subachoque', 'Zipacón'
            ]
            
            for m in municipios_iniciales:
                if not Municipio.query.filter_by(nombre=m).first():
                    db.session.add(Municipio(nombre=m))
            
            admins_default = [
                {'email': 'admin1@uniagraria.edu.co', 'nombre': 'Administrador Principal', 'rol': 'admin'},
                {'email': 'admin2@uniagraria.edu.co', 'nombre': 'Coordinador de Recolección', 'rol': 'admin'},
                {'email': 'admin3@uniagraria.edu.co', 'nombre': 'Director de Proyectos', 'rol': 'admin'},
                {'email': 'admin4@uniagraria.edu.co', 'nombre': 'Supervisor de Campo', 'rol': 'admin'},
                {'email': 'admin5@uniagraria.edu.co', 'nombre': 'Gestor de Calidad', 'rol': 'admin'}
            ]
            
            for admin in admins_default:
                if not Usuario.query.filter_by(email=admin['email']).first():
                    db.session.add(Usuario(**admin))
            
            db.session.commit()
            app.logger.info("✅ Base de datos inicializada correctamente")
            return True
            
        except Exception as e:
            db.session.rollback()
            app.logger.error(f"❌ Error al inicializar base de datos: {str(e)}")
            return False

with app.app_context():
    init_database()

# ============================================================================
# PUNTO DE ENTRADA DE LA APLICACIÓN
# ===========================================================================

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
