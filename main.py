"""
Script principal - Orquestación del proceso PDF a Excel
"""

import os
import sys
from pathlib import Path

# Añadir src al path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from pdf_extractor import PDFExtractor
from excel_filler import ExcelFiller
from intelligent_matcher import IntelligentMatcher
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class PDFToExcelAutomation:
    """Orquestador principal del proceso"""
    
    def __init__(self, pdf_path: str, template_path: str):
        """
        Inicializa el proceso
        
        Args:
            pdf_path: Ruta del PDF a procesar
            template_path: Ruta de la plantilla Excel
        """
        self.pdf_path = pdf_path
        self.template_path = template_path
        self.extracted_data = None
        self.matcher = IntelligentMatcher()
        
    def process(self, output_path: str = None) -> bool:
        """
        Ejecuta el flujo completo: extrae PDF → mapea datos → rellena Excel
        
        Args:
            output_path: Ruta de salida del Excel (opcional)
            
        Returns:
            True si el proceso fue exitoso
        """
        logger.info("=" * 60)
        logger.info("INICIANDO PROCESO PDF → EXCEL AUTOMÁTICO")
        logger.info("=" * 60)
        
        # Paso 1: Validar archivos
        if not self._validate_files():
            return False
        
        # Paso 2: Extraer datos del PDF
        logger.info("\n[1/4] EXTRAYENDO DATOS DEL PDF...")
        if not self._extract_pdf_data():
            return False
        
        # Paso 3: Cargar plantilla Excel
        logger.info("\n[2/4] CARGANDO PLANTILLA EXCEL...")
        if not self._load_excel_template():
            return False
        
        # Paso 4: Mapear datos inteligentemente
        logger.info("\n[3/4] MAPEANDO DATOS INTELIGENTEMENTE...")
        if not self._map_data_intelligently():
            return False
        
        # Paso 5: Rellenar y guardar
        logger.info("\n[4/4] RELLENANDO PLANTILLA Y GUARDANDO...")
        if not self._fill_and_save(output_path):
            return False
        
        logger.info("\n" + "=" * 60)
        logger.info("✓ ¡PROCESO COMPLETADO EXITOSAMENTE!")
        logger.info("=" * 60)
        return True
    
    def _validate_files(self) -> bool:
        """Valida que los archivos existan"""
        if not os.path.exists(self.pdf_path):
            logger.error(f"✗ Archivo PDF no encontrado: {self.pdf_path}")
            return False
        
        if not os.path.exists(self.template_path):
            logger.error(f"✗ Plantilla Excel no encontrada: {self.template_path}")
            return False
        
        logger.info(f"✓ PDF: {self.pdf_path}")
        logger.info(f"✓ Plantilla: {self.template_path}")
        return True
    
    def _extract_pdf_data(self) -> bool:
        """Extrae datos del PDF"""
        try:
            extractor = PDFExtractor(self.pdf_path)
            self.extracted_data = extractor.extract_all()
            
            # Mostrar resumen
            if self.extracted_data.get('structured'):
                logger.info("✓ Datos estructurados extraídos:")
                structured = self.extracted_data['structured']
                for key, values in structured.items():
                    if values:
                        logger.info(f"  - {key}: {len(values)} encontrado(s)")
            
            return True
        except Exception as e:
            logger.error(f"✗ Error al extraer PDF: {e}")
            return False
    
    def _load_excel_template(self) -> bool:
        """Carga la plantilla Excel"""
        try:
            self.excel_filler = ExcelFiller(self.template_path)
            success = self.excel_filler.load_template()
            
            if success:
                # Mostrar campos disponibles
                fields = []
                for row in self.excel_filler.worksheet.iter_rows():
                    for cell in row:
                        if cell.value:
                            fields.append(cell.value)
                
                logger.info(f"✓ Plantilla cargada con {len(set(fields))} campos")
            
            return success
        except Exception as e:
            logger.error(f"✗ Error al cargar plantilla: {e}")
            return False
    
    def _map_data_intelligently(self) -> bool:
        """Mapea datos del PDF a campos de Excel de forma inteligente"""
        try:
            # Obtener campos disponibles
            available_fields = []
            for row in self.excel_filler.worksheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        available_fields.append(cell.value)
            
            available_fields = list(set(available_fields))
            
            # Crear mapeo óptimo
            self.optimal_mapping = self.matcher.create_optimal_mapping(
                self.extracted_data.get('structured', {}),
                available_fields
            )
            
            logger.info(f"✓ Mapeo creado: {len(self.optimal_mapping)} campos mapeados")
            return True
        except Exception as e:
            logger.error(f"✗ Error al mapear datos: {e}")
            return False
    
    def _fill_and_save(self, output_path: str = None) -> bool:
        """Rellena la plantilla y guarda el archivo"""
        try:
            # Rellenar con datos mapeados
            self.excel_filler.fill_with_data(self.optimal_mapping, create_mapping=True)
            
            # Guardar archivo
            success = self.excel_filler.save_file(output_path)
            
            if success and output_path:
                logger.info(f"✓ Archivo guardado en: {output_path}")
            
            self.excel_filler.close()
            return success
        except Exception as e:
            logger.error(f"✗ Error al rellenar y guardar: {e}")
            return False


def main():
    """Función principal"""
    
    # Configuración de ejemplo
    pdf_file = "documentos/ejemplo.pdf"  # Cambiar por tu PDF
    excel_template = "plantillas/plantilla.xlsx"  # Cambiar por tu plantilla
    output_file = "resultados/documento_rellenado.xlsx"
    
    # Crear directorios si no existen
    os.makedirs("documentos", exist_ok=True)
    os.makedirs("plantillas", exist_ok=True)
    os.makedirs("resultados", exist_ok=True)
    
    # Ejecutar proceso
    automation = PDFToExcelAutomation(pdf_file, excel_template)
    success = automation.process(output_file)
    
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
