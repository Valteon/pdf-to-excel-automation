"""
Módulo para matchear datos de PDF con campos Excel de forma inteligente
"""

import re
from typing import Dict, List, Tuple, Any
from difflib import SequenceMatcher
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class IntelligentMatcher:
    """Realiza matching inteligente entre datos extraídos y campos de plantilla"""
    
    def __init__(self, similarity_threshold: float = 0.6):
        """
        Inicializa el matcher
        
        Args:
            similarity_threshold: Umbral de similitud (0-1)
        """
        self.similarity_threshold = similarity_threshold
        self.field_patterns = self._initialize_patterns()
        
    def _initialize_patterns(self) -> Dict[str, List[str]]:
        """
        Inicializa patrones para identificar tipos de campos
        
        Returns:
            Diccionario con patrones por tipo de campo
        """
        return {
            'nombre': ['name', 'nombre', 'apellido', 'full_name', 'person', 'person_name'],
            'email': ['email', 'correo', 'electronic_mail', 'contact_email'],
            'telefono': ['phone', 'telefono', 'movil', 'celular', 'contact_phone'],
            'fecha': ['date', 'fecha', 'dia', 'mes', 'año', 'birth_date', 'fecha_nacimiento'],
            'monto': ['amount', 'monto', 'precio', 'total', 'suma', 'cantidad_dinero'],
            'direccion': ['address', 'direccion', 'domicilio', 'ubicacion', 'street'],
            'empresa': ['company', 'empresa', 'razon_social', 'negocio', 'organization'],
            'documento': ['document', 'documento', 'cedula', 'pasaporte', 'id_number'],
        }
    
    def similarity_ratio(self, a: str, b: str) -> float:
        """
        Calcula la similitud entre dos strings
        
        Args:
            a: Primer string
            b: Segundo string
            
        Returns:
            Ratio de similitud (0-1)
        """
        a_normalized = a.lower().strip().replace('_', ' ')
        b_normalized = b.lower().strip().replace('_', ' ')
        return SequenceMatcher(None, a_normalized, b_normalized).ratio()
    
    def identify_field_type(self, field_name: str) -> str:
        """
        Identifica el tipo de campo basado en su nombre
        
        Args:
            field_name: Nombre del campo
            
        Returns:
            Tipo de campo identificado
        """
        field_name_lower = field_name.lower()
        
        best_match = None
        best_score = 0
        
        for field_type, patterns in self.field_patterns.items():
            for pattern in patterns:
                similarity = self.similarity_ratio(field_name_lower, pattern)
                if similarity > best_score:
                    best_score = similarity
                    best_match = field_type
        
        if best_score >= self.similarity_threshold:
            logger.info(f"✓ Campo '{field_name}' identificado como: {best_match} (score: {best_score:.2f})")
            return best_match
        
        logger.warning(f"⚠ No se pudo identificar tipo de campo: {field_name}")
        return 'unknown'
    
    def match_data_to_fields(self, extracted_data: Dict[str, Any], 
                            available_fields: List[str]) -> Dict[str, Tuple[str, Any]]:
        """
        Intenta hacer matching entre datos extraídos y campos disponibles
        
        Args:
            extracted_data: Datos extraídos del PDF
            available_fields: Campos disponibles en la plantilla
            
        Returns:
            Diccionario con matching {campo_plantilla: (tipo, valor)}
        """
        matches = {}
        
        for field in available_fields:
            field_type = self.identify_field_type(field)
            
            # Buscar datos que coincidan con el tipo identificado
            best_value = None
            best_score = 0
            
            if field_type in extracted_data:
                values = extracted_data[field_type]
                if values:
                    # Para listas, usar el primer elemento
                    best_value = values[0] if isinstance(values, list) else values
                    best_score = 1.0
            
            if best_value:
                matches[field] = (field_type, best_value)
                logger.info(f"✓ {field} → {best_value}")
            else:
                logger.warning(f"⚠ No se encontró valor para {field}")
        
        return matches
    
    def validate_data(self, field_type: str, value: Any) -> Tuple[bool, str]:
        """
        Valida que el valor sea apropiado para el tipo de campo
        
        Args:
            field_type: Tipo de campo
            value: Valor a validar
            
        Returns:
            Tupla (es_válido, mensaje)
        """
        if not value:
            return False, "Valor vacío"
        
        value_str = str(value).strip()
        
        # Validaciones específicas por tipo
        validations = {
            'email': (r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}$', 'Email inválido'),
            'telefono': (r'^\+?[\d\s()-]{9,}$', 'Teléfono inválido'),
            'fecha': (r'^(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})$', 'Fecha inválida'),
            'monto': (r'^[\d,.]+(€|\$|€)?$', 'Monto inválido'),
            'documento': (r'^[A-Za-z0-9-]{5,}$', 'Documento inválido'),
        }
        
        if field_type in validations:
            pattern, error_msg = validations[field_type]
            if re.match(pattern, value_str):
                logger.info(f"✓ Validación exitosa: {field_type} = {value}")
                return True, "Válido"
            else:
                logger.warning(f"⚠ Validación fallida: {field_type} = {value}")
                return False, error_msg
        
        # Si no hay validación específica, considerarlo válido
        return True, "Válido"
    
    def format_value_for_field(self, field_type: str, value: Any) -> Any:
        """
        Formatea el valor según el tipo de campo
        
        Args:
            field_type: Tipo de campo
            value: Valor a formatear
            
        Returns:
            Valor formateado
        """
        value_str = str(value).strip()
        
        if field_type == 'nombre':
            return value_str.title()
        
        elif field_type == 'email':
            return value_str.lower()
        
        elif field_type == 'telefono':
            # Remover caracteres especiales
            cleaned = re.sub(r'[^\d+]', '', value_str)
            return cleaned
        
        elif field_type == 'fecha':
            # Normalizar formato de fecha
            return self._normalize_date(value_str)
        
        elif field_type == 'monto':
            # Remover símbolos y espacios, convertir a número
            cleaned = re.sub(r'[^\d.,]', '', value_str)
            try:
                return float(cleaned.replace(',', ''))
            except:
                return value_str
        
        elif field_type == 'documento':
            return value_str.upper()
        
        return value_str
    
    def _normalize_date(self, date_str: str) -> str:
        """
        Normaliza formato de fecha
        
        Args:
            date_str: String con la fecha
            
        Returns:
            Fecha normalizada (DD/MM/YYYY)
        """
        # Intentar reconocer varios formatos
        patterns = [
            (r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', '{0}/{1}/{2}'),  # DD/MM/YYYY
            (r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', '{2}/{1}/{0}'),  # YYYY/MM/DD
        ]
        
        for pattern, fmt in patterns:
            match = re.match(pattern, date_str)
            if match:
                return fmt.format(*match.groups())
        
        return date_str
    
    def create_optimal_mapping(self, extracted_data: Dict[str, Any],
                              available_fields: List[str],
                              auto_validate: bool = True) -> Dict[str, Any]:
        """
        Crea un mapeo óptimo de datos a campos
        
        Args:
            extracted_data: Datos extraídos
            available_fields: Campos disponibles
            auto_validate: Si se deben validar automáticamente
            
        Returns:
            Diccionario con el mapeo optimizado
        """
        matches = self.match_data_to_fields(extracted_data, available_fields)
        optimal_mapping = {}
        
        for field, (field_type, value) in matches.items():
            if auto_validate:
                is_valid, msg = self.validate_data(field_type, value)
                if not is_valid:
                    logger.warning(f"⚠ Dato inválido para {field}: {msg}")
                    continue
            
            formatted_value = self.format_value_for_field(field_type, value)
            optimal_mapping[field] = formatted_value
        
        logger.info(f"✓ Mapeo óptimo creado con {len(optimal_mapping)} campos")
        return optimal_mapping
