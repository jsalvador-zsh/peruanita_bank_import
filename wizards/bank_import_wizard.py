from odoo import models, fields, api, _
from odoo.exceptions import UserError


class BankImportWizard(models.TransientModel):
    _name = 'bank.import.wizard'
    _description = 'Asistente de Importación Bancaria'

    import_id = fields.Many2one('bank.import', string='Importación', required=True)
    date_from = fields.Date('Fecha Desde', help='Buscar pagos desde esta fecha')
    date_to = fields.Date('Fecha Hasta', help='Buscar pagos hasta esta fecha')
    amount_tolerance = fields.Float('Tolerancia de Monto (%)', default=0.0, 
                                   help='Porcentaje de tolerancia para coincidencia de montos')
    search_in_communication = fields.Boolean('Buscar en Comunicación', default=True)
    search_in_reference = fields.Boolean('Buscar en Referencia', default=True)
    search_in_narration = fields.Boolean('Buscar en Narración', default=True)
    
    def action_advanced_match(self):
        """Realizar búsqueda avanzada de coincidencias"""
        matches_found = 0
        
        # Limpiar matches anteriores
        self.import_id.matched_payment_ids.unlink()
        
        # Configurar dominio de búsqueda de pagos
        payment_domain = [('state', 'in', ['posted', 'sent'])]
        
        if self.date_from:
            payment_domain.append(('date', '>=', self.date_from))
        if self.date_to:
            payment_domain.append(('date', '<=', self.date_to))
        
        payments = self.env['account.payment'].search(payment_domain)
        
        for line in self.import_id.line_ids:
            matched_payments = self._find_advanced_matches(line, payments)
            matches_found += len(matched_payments)
        
        if matches_found > 0:
            self.import_id.state = 'matched'
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': _('Éxito'),
                    'message': _('Se encontraron %d coincidencias.') % matches_found,
                    'type': 'success',
                    'sticky': False,
                }
            }
        else:
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': _('Sin resultados'),
                    'message': _('No se encontraron coincidencias con los criterios especificados.'),
                    'type': 'warning',
                    'sticky': False,
                }
            }

    def _find_advanced_matches(self, import_line, payments):
        """Buscar coincidencias avanzadas para una línea"""
        matched_payments = []
        
        for payment in payments:
            match_score = self._calculate_match_score(import_line, payment)
            
            if match_score > 0:
                # Crear el match
                self.env['bank.import.match'].create({
                    'import_id': self.import_id.id,
                    'import_line_id': import_line.id,
                    'payment_id': payment.id,
                    'match_type': 'exact' if match_score >= 100 else 'partial'
                })
                matched_payments.append(payment)
        
        return matched_payments

    def _calculate_match_score(self, import_line, payment):
        """Calcular puntaje de coincidencia entre línea de importación y pago"""
        score = 0
        
        # Verificar monto con tolerancia
        if self._amounts_match(import_line.amount, payment.amount):
            score += 50
        
        # Verificar número de operación
        if import_line.operation_number and self._operation_number_matches(import_line.operation_number, payment):
            score += 50
        
        return score

    def _amounts_match(self, import_amount, payment_amount):
        """Verificar si los montos coinciden considerando tolerancia"""
        if self.amount_tolerance == 0.0:
            return abs(import_amount) == payment_amount
        
        tolerance = abs(import_amount) * (self.amount_tolerance / 100.0)
        return abs(abs(import_amount) - payment_amount) <= tolerance

    def _operation_number_matches(self, operation_number, payment):
        """Verificar si el número de operación coincide en algún campo del pago"""
        fields_to_check = []
        
        if self.search_in_reference and payment.name:
            fields_to_check.append(payment.name)
        if self.search_in_communication and payment.communication:
            fields_to_check.append(payment.communication)
        
        # Verificar campos adicionales si existen
        if self.search_in_narration:
            if hasattr(payment, 'memo') and payment.memo:
                fields_to_check.append(payment.memo)
            if hasattr(payment, 'narration') and payment.narration:
                fields_to_check.append(payment.narration)
        
        for field_value in fields_to_check:
            if operation_number in str(field_value):
                return True
        
        return False