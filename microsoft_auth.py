import os



"""
Microsoft 365 Authentication Module - UNIAGRARIA
Sistema de autenticación integrado con Azure AD
"""

import os
import logging
import importlib.util
import subprocess
import sys
from flask import session, redirect, url_for, flash
from functools import wraps

logger = logging.getLogger(__name__)

# ============================================================
# ✅ Detección e instalación automática de MSAL y Requests
# ============================================================

def get_db_connection():
    '''Obtiene conexión a la base de datos'''
    try:
        import psycopg2
        import os
        from flask import current_app
        
        # Obtener DATABASE_URL de múltiples fuentes
        db_url = os.getenv('DATABASE_URL')
        if not db_url and current_app:
            db_url = current_app.config.get('DATABASE_URL')
        
        if not db_url:
            raise ValueError('DATABASE_URL no configurada')
        
        # Crear conexión
        conn = psycopg2.connect(db_url)
        return conn
        
    except ImportError:
        # Intentar instalar psycopg2-binary
        import sys
        import subprocess
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'psycopg2-binary'])
        import psycopg2
        import os
        return psycopg2.connect(os.getenv('DATABASE_URL'))
    except Exception as e:
        print(f'❌ Error en get_db_connection: {e}')
        raise
def get_db_connection():
    '''Obtiene conexión a la base de datos'''
    import psycopg2
    import os
    return psycopg2.connect(os.getenv('DATABASE_URL'))

def ensure_msal_installed():
    """Garantiza que las librerías msal y requests estén disponibles."""
    try:
        msal_spec = importlib.util.find_spec("msal")
        requests_spec = importlib.util.find_spec("requests")

        if msal_spec and requests_spec:
            import msal, requests  # noqa
            return True

        logger.warning("⚠️ MSAL o Requests no detectadas, instalando...")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "msal", "requests"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        import msal, requests  # noqa
        logger.info("✅ Librerías MSAL y Requests instaladas correctamente.")
        return True

    except Exception as e:
        logger.error(f"❌ No se pudieron cargar MSAL/Requests: {e}")
        return False


MSAL_AVAILABLE = ensure_msal_installed()

# ============================================================
# 🔹 Clase de configuración de Microsoft Azure AD
# ============================================================
class MicrosoftConfig:
    def __init__(self):
        self.client_id = os.getenv("AZURE_CLIENT_ID")
        self.client_secret = os.getenv("AZURE_CLIENT_SECRET")
        self.tenant_id = os.getenv("AZURE_TENANT_ID")
        self.redirect_uri = os.getenv(
            "AZURE_REDIRECT_URI", "http://localhost:5000/auth/microsoft/callback"
        )
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id or 'common'}"
        self.scopes = ["User.Read"]
        self.graph_endpoint = "https://graph.microsoft.com/v1.0"

    def is_configured(self):
        return all([
            self.client_id, 
            self.client_secret, 
            self.tenant_id, 
            MSAL_AVAILABLE
        ])

    def get_status(self):
        if not MSAL_AVAILABLE:
            return "ERROR: pip install msal requests"
        if not self.client_id:
            return "PENDIENTE: AZURE_CLIENT_ID"
        if not self.client_secret:
            return "PENDIENTE: AZURE_CLIENT_SECRET"
        if not self.tenant_id:
            return "PENDIENTE: AZURE_TENANT_ID"
        return "CONFIGURADO ✅"


ms_config = MicrosoftConfig()


# ============================================================
# 🔹 Funciones principales de autenticación Microsoft
# ============================================================
def is_microsoft_enabled():
    """Verifica si Microsoft Auth está habilitado"""
    return ms_config.is_configured()


def get_msal_app():
    """Obtiene instancia de MSAL"""
    if not ms_config.is_configured():
        return None
    try:
        import msal
        return msal.ConfidentialClientApplication(
            ms_config.client_id,
            authority=ms_config.authority,
            client_credential=ms_config.client_secret
        )
    except Exception as e:
        logger.error(f"❌ Error MSAL: {e}")
        return None


def get_auth_url():
    """Genera URL de autorización Microsoft"""
    app = get_msal_app()
    if not app:
        return None
    try:
        return app.get_authorization_request_url(
            scopes=ms_config.scopes,
            redirect_uri=ms_config.redirect_uri,
            state=os.urandom(16).hex()
        )
    except Exception as e:
        logger.error(f"❌ Error generando auth URL: {e}")
        return None


def get_token_from_code(code):
    """Intercambia código por token de acceso"""
    app = get_msal_app()
    if not app:
        return None
    try:
        return app.acquire_token_by_authorization_code(
            code,
            scopes=ms_config.scopes,
            redirect_uri=ms_config.redirect_uri
        )
    except Exception as e:
        logger.error(f"❌ Error obteniendo token: {e}")
        return None


def get_user_info(access_token):
    """Obtiene información del usuario desde Microsoft Graph"""
    if not access_token:
        print("❌ Token de acceso vacío")
        return None
    
    import requests
    
    headers = {"Authorization": f"Bearer {access_token}"}
    
    try:
        response = requests.get(
            f"{ms_config.graph_endpoint}/me", 
            headers=headers, 
            timeout=10
        )
        
        print(f"📊 Respuesta de Microsoft Graph: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            
            # 📝 DEBUG COMPLETO - Mostrar TODOS los datos
            print("\n" + "="*60)
            print("🔍 DATOS COMPLETOS DEVUELTOS POR MICROSOFT GRAPH:")
            print("="*60)
            for key, value in data.items():
                print(f"  {key}: {value}")
            print("="*60)
            
            # Intentar obtener email de diferentes campos
            email = None
            
            # Posibles campos donde puede estar el email
            possible_email_fields = ['mail', 'userPrincipalName', 'email', 'primaryEmail']
            
            for field in possible_email_fields:
                if field in data and data[field]:
                    email = data[field].lower()
                    print(f"  📧 Email encontrado en campo '{field}': {email}")
                    break
            
            if not email:
                print("❌ No se encontró email en los datos de Graph")
                return None
            
            return {
                "id": data.get("id"),
                "email": email,
                "name": data.get("displayName", "Usuario"),
                "given_name": data.get("givenName"),
                "surname": data.get("surname"),
                "raw_data": data  # Para debugging
            }
        else:
            print(f"❌ Error de Graph API: {response.status_code}")
            print(f"   Respuesta: {response.text}")
            return None
            
    except Exception as e:
        print(f"❌ Error obteniendo user info: {e}")
        import traceback
        traceback.print_exc()
        return None

# ============================================================
# 🔹 Validación de correos institucionales UNIAGRARIA
# ============================================================
def validate_uniagraria_email(email):
    """
    🔍 VALIDACIÓN MODIFICADA - PERMITE EMAIL DE LA COORDINADORA
    Solo permite correos @uniagraria.edu.co
    """
    if not email:
        print("❌ Email vacío")
        return False
    
    # Convertir a string y limpiar
    email_str = str(email).strip().lower()
    
    print(f"📧 Validando email: {email_str}")
    
    # 📌 LISTA DE EMAILS ESPECIALES PERMITIDOS
    emails_especiales_permitidos = [
        'colorador.yeimy@uniagraria.edu.co',  # COORDINADORA ADMINISTRATIVA
        'admin@uniagraria.edu.co',            # ADMIN PRINCIPAL
        'admin_backup@uniagraria.edu.co',     # ADMIN DE RESPALDO
        'operador@uniagraria.edu.co',         # OPERADOR
        # Agrega más emails especiales aquí si es necesario
    ]
    
    # 1. PRIMERO verificar si es un email especial permitido
    if email_str in emails_especiales_permitidos:
        print(f"✅✅✅ EMAIL ESPECIAL PERMITIDO: {email_str}")
        print(f"   Esta es la COORDINADORA ADMINISTRATIVA")
        return True
    
    # 2. LUEGO verificar dominio @uniagraria.edu.co
    allowed_domains = ["uniagraria.edu.co"]
    
    for domain in allowed_domains:
        if email_str.endswith("@" + domain):
            print(f"✅ Email válido: Dominio {domain} permitido")
            return True
    
    # Si llega aquí, el email no está permitido
    print(f"❌ Email rechazado: No es @uniagraria.edu.co ni está en lista especial")
    print(f"   Email recibido: {email_str}")
    print(f"   Emails especiales permitidos: {emails_especiales_permitidos}")
    print(f"   Dominios permitidos: {allowed_domains}")
    return False
def create_or_update_user(user_info):
    """
    🔍 DEBUG EXTENDIDO - Crear o actualizar usuario después del login de Microsoft
    """
    print("\n" + "="*80)
    print("🔍 DEBUG: create_or_update_user - INICIO COMPLETO")
    print("="*80)
    
    # 1. MOSTRAR TODA LA INFORMACIÓN RECIBIDA
    print("📋 user_info COMPLETO RECIBIDO:")
    print("-"*40)
    for key, value in user_info.items():
        if key == 'raw_data' and isinstance(value, dict):
            print(f"   {key}: (dict con {len(value)} campos)")
        else:
            print(f"   {key}: {value}")
    
    # 2. EXTRAER DATOS PRINCIPALES
    email = user_info.get("email")
    microsoft_id = user_info.get("id")
    nombre_completo = user_info.get("name", "Usuario Microsoft")
    
    print(f"\n🎯 DATOS PRINCIPALES EXTRAÍDOS:")
    print(f"   📧 Email: {email}")
    print(f"   🆔 Microsoft ID: {microsoft_id}")
    print(f"   👤 Nombre completo: {nombre_completo}")
    
    # 3. VALIDACIÓN CRÍTICA
    if not email:
        print("\n❌ ERROR CRÍTICO: user_info NO contiene campo 'email'")
        return None
    
    # 4. VALIDAR EMAIL
    print(f"\n🔐 VALIDANDO EMAIL...")
    if not validate_uniagraria_email(email):
        print(f"❌ Email rechazado en validación: {email}")
        return None
    
    # 5. LIMPIAR Y PREPARAR DATOS
    email_limpio = str(email).lower().strip()
    username = email_limpio.split('@')[0]  # Extraer username del email
    
    print(f"\n🧹 DATOS LIMPIOS:")
    print(f"   Email limpio: {email_limpio}")
    print(f"   Username: {username}")
    print(f"   Microsoft ID: {microsoft_id}")
    
    # 6. CONEXIÓN A BASE DE DATOS
    print(f"\n🗄️  CONECTANDO A BASE DE DATOS...")
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        print(f"✅ Conexión a BD establecida")
        
        # 7. BUSCAR USUARIO EXISTENTE (CORREGIDO - usar email y microsoft_id)
        print(f"\n🔍 BUSCANDO USUARIO EXISTENTE...")
        
        # 🔥 CAMBIAR ESTA CONSULTA:
        # Versión INCORRECTA: SELECT id, nombre, email, rol, activo, microsoft_id
        # Versión CORRECTA: SELECT id, username, email, rol, activo, microsoft_id
        
        cursor.execute("""
            SELECT id, username, email, rol, activo, microsoft_id
            FROM usuarios 
            WHERE email = %s OR microsoft_id = %s
        """, (email_limpio, microsoft_id))
        
        existing = cursor.fetchone()
        
        if existing:
            # 8A. USUARIO EXISTE - ACTUALIZAR
            user_id = existing[0]
            username_actual = existing[1]
            email_actual = existing[2]
            rol = existing[3]
            activo = bool(existing[4])
            ms_id_existente = existing[5]
            
            print(f"📝 USUARIO EXISTENTE ENCONTRADO:")
            print(f"   ID: {user_id}")
            print(f"   Username actual: {username_actual}")
            print(f"   Email actual: {email_actual}")
            print(f"   Rol: {rol}")
            print(f"   Activo: {'✅' if activo else '❌'}")
            print(f"   Microsoft ID en BD: {ms_id_existente}")
            
            if not activo:
                print(f"❌ USUARIO INACTIVO: {email_limpio}")
                cursor.close()
                conn.close()
                return None
            
            # Determinar si necesita actualización
            necesita_actualizacion = False
            cambios = []
            
            if not ms_id_existente or ms_id_existente != microsoft_id:
                necesita_actualizacion = True
                cambios.append(f"Microsoft ID: {ms_id_existente} → {microsoft_id}")
            
            if username_actual != username:
                necesita_actualizacion = True
                cambios.append(f"Username: {username_actual} → {username}")
            
            if necesita_actualizacion:
                print(f"🔄 ACTUALIZANDO USUARIO...")
                print(f"   Cambios: {', '.join(cambios)}")
                
                cursor.execute("""
                    UPDATE usuarios 
                    SET microsoft_id = %s, 
                        username = %s,
                        nombre_completo = %s,
                        ultimo_acceso = CURRENT_TIMESTAMP
                    WHERE id = %s
                """, (microsoft_id, username, nombre_completo, user_id))
                
                conn.commit()
                print(f"✅ USUARIO ACTUALIZADO: {email_limpio}")
            else:
                print(f"ℹ️  Usuario ya actualizado, solo registro de acceso")
                cursor.execute("""
                    UPDATE usuarios 
                    SET ultimo_acceso = CURRENT_TIMESTAMP
                    WHERE id = %s
                """, (user_id,))
                conn.commit()
            
        else:
            # 8B. USUARIO NUEVO - CREAR (CORREGIDO)
            print(f"➕ CREANDO NUEVO USUARIO...")
            print(f"   Email: {email_limpio}")
            print(f"   Username: {username}")
            print(f"   Nombre completo: {nombre_completo}")
            print(f"   Microsoft ID: {microsoft_id}")
            
            # Determinar rol basado en email
            rol_asignado = "usuario"  # Por defecto
            if "admin" in email_limpio or "administrador" in email_limpio.lower():
                rol_asignado = "admin"
            elif "colorador.yeimy" in email_limpio:  # Específico para tu jefa
                rol_asignado = "admin"
            elif "profesor" in email_limpio or "docente" in email_limpio.lower():
                rol_asignado = "profesor"
            
            print(f"   Rol asignado: {rol_asignado}")
            
            # 🔥 CAMBIAR ESTA INSERCIÓN:
            # Usar las columnas CORRECTAS según tu estructura SQL
            
            cursor.execute("""
                INSERT INTO usuarios (
                    username, 
                    email, 
                    nombre_completo,
                    microsoft_id, 
                    rol, 
                    activo,
                    fecha_creacion,
                    ultimo_acceso
                ) VALUES (%s, %s, %s, %s, %s, TRUE, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                RETURNING id, rol
            """, (
                username,           # username (NO nombre)
                email_limpio,       # email
                nombre_completo,    # nombre_completo
                microsoft_id,       # microsoft_id
                rol_asignado        # rol
            ))
            
            result = cursor.fetchone()
            user_id = result[0]
            rol_final = result[1]
            
            conn.commit()
            print(f"✅ NUEVO USUARIO CREADO:")
            print(f"   ID: {user_id}")
            print(f"   Email: {email_limpio}")
            print(f"   Username: {username}")
            print(f"   Rol: {rol_final}")
            print(f"   Microsoft ID: {microsoft_id}")
        
        cursor.close()
        conn.close()
        
        # 9. RETORNAR DATOS DEL USUARIO
        usuario_final = {
            "id": user_id,
            "email": email_limpio,
            "name": nombre_completo,
            "rol": rol if 'rol' in locals() else rol_final,
            "microsoft_id": microsoft_id
        }
        
        print(f"\n🎉 USUARIO FINAL PREPARADO:")
        for key, value in usuario_final.items():
            print(f"   {key}: {value}")
        
        print("\n" + "="*80)
        print("🔍 DEBUG: create_or_update_user - FIN EXITOSO")
        print("="*80)
        
        return usuario_final
        
    except Exception as e:
        print(f"\n❌ ERROR CRÍTICO EN create_or_update_user:")
        print(f"   Tipo: {type(e).__name__}")
        print(f"   Mensaje: {str(e)}")
        
        import traceback
        print(f"\n📜 TRACEBACK COMPLETO:")
        traceback.print_exc()
        
        return None


# ============================================================
# 🔹 Decoradores de acceso (compatibles con auth.py)
# ============================================================
def login_required(f):
    """Decorador para rutas que requieren autenticación"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            flash("⚠️ Debes iniciar sesión para acceder", "warning")
            return redirect(url_for("auth_bp.login"))
        return f(*args, **kwargs)
    return decorated_function


def admin_required(f):
    """Decorador para rutas que requieren rol admin"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            flash("⚠️ Debes iniciar sesión", "warning")
            return redirect(url_for("auth_bp.login"))
        
        if session.get("user_rol") != "admin":
            flash("❌ No tienes permisos de administrador", "danger")
            return redirect(url_for("routes_bp.activos_list"))
        
        return f(*args, **kwargs)
    return decorated_function


# ========================================
