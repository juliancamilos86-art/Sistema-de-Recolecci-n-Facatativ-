import os
import sys
import psycopg2
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT
from dotenv import load_dotenv

load_dotenv()

def init_database():
    """Inicializar base de datos en NeonTech PostgreSQL"""
    try:
        database_url = os.environ.get('DATABASE_URL')
        if not database_url:
            print("❌ ERROR: DATABASE_URL no está configurada en .env")
            sys.exit(1)
        
        conn = psycopg2.connect(database_url)
        conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
        cursor = conn.cursor()
        
        print("✅ Conectado a NeonTech PostgreSQL")
        
        # Crear tablas
        tables_sql = """
        -- Tabla de usuarios
        CREATE TABLE IF NOT EXISTS usuarios (
            id SERIAL PRIMARY KEY,
            email VARCHAR(255) UNIQUE NOT NULL,
            nombre VARCHAR(255) NOT NULL,
            apellido VARCHAR(255),
            azure_id VARCHAR(255) UNIQUE,
            rol VARCHAR(50) DEFAULT 'usuario',
            avatar_url VARCHAR(500),
            activo BOOLEAN DEFAULT true,
            ultimo_acceso TIMESTAMP,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        
        -- Tabla de municipios
        CREATE TABLE IF NOT EXISTS municipios (
            id SERIAL PRIMARY KEY,
            nombre VARCHAR(100) UNIQUE NOT NULL,
            departamento VARCHAR(100) DEFAULT 'Cundinamarca',
            activo BOOLEAN DEFAULT true
        );
        
        -- Tabla de instituciones
        CREATE TABLE IF NOT EXISTS instituciones (
            id SERIAL PRIMARY KEY,
            nombre VARCHAR(255) NOT NULL,
            municipio_id INTEGER REFERENCES municipios(id),
            direccion VARCHAR(255),
            telefono VARCHAR(50),
            contacto VARCHAR(255),
            activo BOOLEAN DEFAULT true
        );
        
        -- Tabla principal de recolección
        CREATE TABLE IF NOT EXISTS recoleccion_datos (
            id SERIAL PRIMARY KEY,
            ficha_toma_registro VARCHAR(50),
            asesor VARCHAR(255),
            municipio_id INTEGER REFERENCES municipios(id),
            institucion_id INTEGER REFERENCES instituciones(id),
            realizador_nombre VARCHAR(255),
            realizador_apellidos VARCHAR(255),
            matricula_documento VARCHAR(100),
            telefono VARCHAR(50),
            correo VARCHAR(255),
            grado VARCHAR(50),
            instalacion_educativa VARCHAR(255),
            pista VARCHAR(100),
            programa_interes VARCHAR(255),
            jornada_interes VARCHAR(100),
            pendido_interes TEXT,
            asesoria_migradora TEXT,
            estado VARCHAR(50) DEFAULT 'pendiente',
            ano_periodo VARCHAR(50) DEFAULT '2026-1',
            observacion TEXT,
            usuario_registro_id INTEGER REFERENCES usuarios(id),
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            fecha_actualizacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        
        -- Tabla de ferias
        CREATE TABLE IF NOT EXISTS ferias (
            id SERIAL PRIMARY KEY,
            nombre VARCHAR(255) NOT NULL,
            fecha_inicio DATE,
            fecha_fin DATE,
            ubicacion VARCHAR(255),
            municipio_id INTEGER REFERENCES municipios(id),
            descripcion TEXT,
            activa BOOLEAN DEFAULT true
        );
        
        -- Tabla de imágenes de ferias
        CREATE TABLE IF NOT EXISTS ferias_imagenes (
            id SERIAL PRIMARY KEY,
            feria_id INTEGER REFERENCES ferias(id),
            public_id VARCHAR(255),
            url VARCHAR(500),
            descripcion TEXT,
            usuario_subida_id INTEGER REFERENCES usuarios(id),
            fecha_subida TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        
        -- Tabla de archivos importados
        CREATE TABLE IF NOT EXISTS archivos_importados (
            id SERIAL PRIMARY KEY,
            nombre_archivo VARCHAR(255),
            tipo VARCHAR(50),
            url VARCHAR(500),
            usuario_importo_id INTEGER REFERENCES usuarios(id),
            fecha_importacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            registros_procesados INTEGER DEFAULT 0,
            estado VARCHAR(50) DEFAULT 'completado',
            metadata JSONB
        );
        
        -- Índices para mejorar rendimiento
        CREATE INDEX IF NOT EXISTS idx_recoleccion_municipio ON recoleccion_datos(municipio_id);
        CREATE INDEX IF NOT EXISTS idx_recoleccion_fecha ON recoleccion_datos(fecha_registro);
        CREATE INDEX IF NOT EXISTS idx_recoleccion_estado ON recoleccion_datos(estado);
        CREATE INDEX IF NOT EXISTS idx_recoleccion_periodo ON recoleccion_datos(ano_periodo);
        CREATE INDEX IF NOT EXISTS idx_usuarios_email ON usuarios(email);
        CREATE INDEX IF NOT EXISTS idx_ferias_fechas ON ferias(fecha_inicio, fecha_fin);
        """
        
        cursor.execute(tables_sql)
        print("✅ Tablas creadas/verificadas exitosamente")
        
        # Insertar municipios iniciales
        municipios = [
            'Facatativá', 'Bogotá', 'Madrid', 'Mosquera', 
            'Funza', 'El Rosal', 'Subachoque', 'Zipacón'
        ]
        
        for m in municipios:
            cursor.execute(
                "INSERT INTO municipios (nombre) VALUES (%s) ON CONFLICT (nombre) DO NOTHING",
                (m,)
            )
        
        # Insertar 5 usuarios admin por defecto
        admins = [
            ('admin1@uniagraria.edu.co', 'Administrador Principal', 'admin'),
            ('admin2@uniagraria.edu.co', 'Coordinador de Recolección', 'admin'),
            ('admin3@uniagraria.edu.co', 'Director de Proyectos', 'admin'),
            ('admin4@uniagraria.edu.co', 'Supervisor de Campo', 'admin'),
            ('admin5@uniagraria.edu.co', 'Gestor de Calidad', 'admin')
        ]
        
        for email, nombre, rol in admins:
            cursor.execute(
                """
                INSERT INTO usuarios (email, nombre, rol) 
                VALUES (%s, %s, %s) 
                ON CONFLICT (email) DO NOTHING
                """,
                (email, nombre, rol)
            )
        
        conn.commit()
        print("✅ Datos iniciales insertados correctamente")
        
        cursor.close()
        conn.close()
        
        print("\n" + "="*50)
        print("✅ BASE DE DATOS INICIALIZADA CON ÉXITO")
        print("="*50)
        print("\n📊 NeonTech PostgreSQL configurado")
        print("👥 5 usuarios admin creados por defecto")
        print("🏛️ Municipios de Cundinamarca agregados")
        print("📱 Listo para producción con Azure AD y Cloudinary")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    print("🚀 Inicializando base de datos para UNIAGRARIA - Facatativá 2026...")
    print("="*50)
    init_database()
