# -*- coding: utf-8 -*-
"""
Script de Processamento em Batch
================================
Executa o processamento de todos os relatorios disponiveis sem interface grafica.
"""

import os
import sys
import logging
import time
from datetime import datetime

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('batch_processing.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# Importar classes do main.py
from main import (
    CompanyProcessor,
    RankingProcessor,
    RankingIntegracaoProcessor,
    RankingOuroMedianoProcessor,
    RankingKmProporcionalProcessor,
    TurnosIntegracaoProcessor,
    ResumoMotoristaClienteProcessor,
    normalize_matricula
)

class BatchProcessor:
    """Processador em batch para todos os relatorios"""
    
    def __init__(self, base_dir: str, output_dir: str, version_suffix: str = ""):
        self.base_dir = base_dir
        self.output_dir = output_dir
        self.version_suffix = version_suffix
        
        # Inicializar processadores
        self.company_processor = CompanyProcessor(base_dir, output_dir, version_suffix)
        self.ranking_processor = RankingProcessor(base_dir, output_dir, version_suffix)
        self.ranking_integracao_processor = RankingIntegracaoProcessor(base_dir, output_dir, version_suffix)
        self.ranking_ouro_mediano_processor = RankingOuroMedianoProcessor(base_dir, output_dir, version_suffix)
        self.ranking_km_proporcional_processor = RankingKmProporcionalProcessor(base_dir, output_dir, version_suffix)
        self.turnos_integracao_processor = TurnosIntegracaoProcessor(base_dir, output_dir, version_suffix)
        
        # Estatisticas
        self.stats = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'skipped': 0
        }
    
    def process_abst_mot_por_empresa(self, company: str, period: str) -> bool:
        """Processa Abst_Mot_Por_empresa para uma empresa e periodo"""
        try:
            logging.info(f"[Abst_Mot_Por_empresa] Processando {company} - {period}")
            
            files = self.company_processor.get_company_files(company)
            target_file = None
            
            for f in files:
                if f['month_year'] == period:
                    target_file = f
                    break
            
            if not target_file:
                logging.warning(f"[Abst_Mot_Por_empresa] Arquivos nao encontrados para {company} - {period}")
                return False
            
            result = self.company_processor.process_company_files(
                target_file['supply'],
                target_file['drivers'],
                company,
                period
            )
            
            if result is not None:
                logging.info(f"[Abst_Mot_Por_empresa] Sucesso: {company} - {period}")
                return True
            else:
                logging.error(f"[Abst_Mot_Por_empresa] Falha: {company} - {period}")
                return False
                
        except Exception as e:
            logging.error(f"[Abst_Mot_Por_empresa] Erro {company} - {period}: {str(e)}")
            return False
    
    def process_ranking_por_empresa(self, company: str, period: str) -> bool:
        """Processa Ranking_Por_Empresa para uma empresa e periodo"""
        try:
            logging.info(f"[Ranking_Por_Empresa] Processando {company} - {period}")
            
            df_result = self.ranking_processor.process_company_period(company, period)
            
            if df_result is not None:
                output_file = self.ranking_processor.create_report(df_result, company, period)
                if output_file:
                    logging.info(f"[Ranking_Por_Empresa] Sucesso: {company} - {period}")
                    return True
            
            logging.error(f"[Ranking_Por_Empresa] Falha: {company} - {period}")
            return False
            
        except Exception as e:
            logging.error(f"[Ranking_Por_Empresa] Erro {company} - {period}: {str(e)}")
            return False
    
    def process_ranking_integracao(self, company: str, period: str) -> bool:
        """Processa Ranking_Integracao para uma empresa e periodo"""
        try:
            logging.info(f"[Ranking_Integracao] Processando {company} - {period}")
            
            df_result = self.ranking_integracao_processor.process_company_period(company, period)
            
            if df_result is not None:
                output_file = self.ranking_integracao_processor.create_report(df_result, company, period)
                if output_file:
                    logging.info(f"[Ranking_Integracao] Sucesso: {company} - {period}")
                    return True
            
            logging.error(f"[Ranking_Integracao] Falha: {company} - {period}")
            return False
            
        except Exception as e:
            logging.error(f"[Ranking_Integracao] Erro {company} - {period}: {str(e)}")
            return False
    
    def process_turnos_integracao(self, company: str, period: str) -> bool:
        """Processa Turnos_Integracao para uma empresa e periodo"""
        try:
            logging.info(f"[Turnos_Integracao] Processando {company} - {period}")
            
            df_result = self.turnos_integracao_processor.process_company_period(company, period)
            
            if df_result is not None:
                output_file = self.turnos_integracao_processor.create_report(df_result, company, period)
                if output_file:
                    logging.info(f"[Turnos_Integracao] Sucesso: {company} - {period}")
                    return True
            
            logging.error(f"[Turnos_Integracao] Falha: {company} - {period}")
            return False
            
        except Exception as e:
            logging.error(f"[Turnos_Integracao] Erro {company} - {period}: {str(e)}")
            return False
    
    def process_ranking_km_proporcional(self, company: str, period: str) -> bool:
        """Processa Ranking_Km_Proporcional para uma empresa e periodo"""
        try:
            logging.info(f"[Ranking_Km_Proporcional] Processando {company} - {period}")
            
            result = self.ranking_km_proporcional_processor.process_company_period(company, period)
            
            if result is not None:
                logging.info(f"[Ranking_Km_Proporcional] Sucesso: {company} - {period}")
                return True
            
            logging.error(f"[Ranking_Km_Proporcional] Falha: {company} - {period}")
            return False
            
        except Exception as e:
            logging.error(f"[Ranking_Km_Proporcional] Erro {company} - {period}: {str(e)}")
            return False
    
    def process_ranking_ouro_mediano(self, companies: list = None, periods: list = None) -> bool:
        """Processa Ranking_Ouro_Mediano (consolidacao)"""
        try:
            logging.info(f"[Ranking_Ouro_Mediano] Processando consolidacao...")
            
            df_result = self.ranking_ouro_mediano_processor.process_consolidation(companies, periods)
            
            if df_result is not None and not df_result.empty:
                output_file = self.ranking_ouro_mediano_processor.create_consolidated_report(df_result, periods, companies)
                if output_file:
                    logging.info(f"[Ranking_Ouro_Mediano] Sucesso: {len(df_result)} registros consolidados")
                    return True
            
            logging.warning(f"[Ranking_Ouro_Mediano] Nenhum dado para consolidar")
            return False
            
        except Exception as e:
            logging.error(f"[Ranking_Ouro_Mediano] Erro: {str(e)}")
            return False
    
    def check_abst_mot_exists(self, company: str, period: str) -> bool:
        """Verifica se Abst_Mot_Por_empresa ja existe"""
        month, year = period.split("_")
        path = os.path.join(
            self.output_dir, "Abst_Mot_Por_empresa",
            company, year, month.zfill(2)
        )
        if os.path.exists(path):
            for f in os.listdir(path):
                if f.startswith("Detalhado_") and f.endswith(".xlsx"):
                    return True
        return False
    
    def check_ranking_por_empresa_exists(self, company: str, period: str) -> bool:
        """Verifica se Ranking_Por_Empresa ja existe"""
        month, year = period.split("_")
        path = os.path.join(
            self.output_dir, "Ranking_Por_Empresa",
            company, year, month.zfill(2)
        )
        if os.path.exists(path):
            for f in os.listdir(path):
                if f.startswith("Ranking_Por_Empresa_") and f.endswith(".xlsx"):
                    return True
        return False
    
    def run_all(self):
        """Executa todo o processamento em ordem"""
        start_time = time.time()
        
        logging.info("=" * 80)
        logging.info("INICIANDO PROCESSAMENTO EM BATCH")
        logging.info(f"Base: {self.base_dir}")
        logging.info(f"Saida: {self.output_dir}")
        logging.info("=" * 80)
        
        # FASE 1: Processar Abst_Mot_Por_empresa (base para os demais)
        logging.info("\n" + "=" * 80)
        logging.info("FASE 1: Abst_Mot_Por_empresa")
        logging.info("=" * 80)
        
        companies_abst = self.company_processor.find_available_companies()
        logging.info(f"Empresas disponiveis: {len(companies_abst)}")
        
        abst_processed = []
        for company in companies_abst:
            files = self.company_processor.get_company_files(company)
            for f in files:
                period = f['month_year']
                self.stats['total'] += 1
                
                if self.check_abst_mot_exists(company, period):
                    logging.info(f"[SKIP] Abst_Mot_Por_empresa {company} - {period} ja existe")
                    self.stats['skipped'] += 1
                    abst_processed.append((company, period))
                else:
                    if self.process_abst_mot_por_empresa(company, period):
                        self.stats['success'] += 1
                        abst_processed.append((company, period))
                    else:
                        self.stats['failed'] += 1
        
        # FASE 2: Processar Ranking_Por_Empresa (precisa de arquivos Ranking e Turnos_128)
        logging.info("\n" + "=" * 80)
        logging.info("FASE 2: Ranking_Por_Empresa")
        logging.info("=" * 80)
        
        companies_ranking = self.ranking_processor.find_available_companies()
        logging.info(f"Empresas disponiveis para Ranking: {len(companies_ranking)}")
        
        ranking_processed = []
        for company in companies_ranking:
            periods = self.ranking_processor.find_available_periods(company)
            for period in periods:
                self.stats['total'] += 1
                
                if self.check_ranking_por_empresa_exists(company, period):
                    logging.info(f"[SKIP] Ranking_Por_Empresa {company} - {period} ja existe")
                    self.stats['skipped'] += 1
                    ranking_processed.append((company, period))
                else:
                    if self.process_ranking_por_empresa(company, period):
                        self.stats['success'] += 1
                        ranking_processed.append((company, period))
                    else:
                        self.stats['failed'] += 1
        
        # FASE 3: Processar Ranking_Integracao (precisa de Ranking, Turnos_128 e Abst_Mot)
        logging.info("\n" + "=" * 80)
        logging.info("FASE 3: Ranking_Integracao")
        logging.info("=" * 80)
        
        companies_integracao = self.ranking_integracao_processor.find_available_companies()
        logging.info(f"Empresas disponiveis para Ranking_Integracao: {len(companies_integracao)}")
        
        for company in companies_integracao:
            periods = self.ranking_integracao_processor.find_available_periods(company)
            for period in periods:
                # Verificar se Abst_Mot existe
                if not self.check_abst_mot_exists(company, period):
                    logging.warning(f"[SKIP] Ranking_Integracao {company} - {period}: Abst_Mot nao existe")
                    continue
                
                self.stats['total'] += 1
                if self.process_ranking_integracao(company, period):
                    self.stats['success'] += 1
                else:
                    self.stats['failed'] += 1
        
        # FASE 4: Processar Ranking_Km_Proporcional (precisa de Abst_Mot/Detalhado)
        logging.info("\n" + "=" * 80)
        logging.info("FASE 4: Ranking_Km_Proporcional")
        logging.info("=" * 80)
        
        companies_km_prop = self.ranking_km_proporcional_processor.find_available_companies()
        logging.info(f"Empresas disponiveis para Ranking_Km_Proporcional: {len(companies_km_prop)}")
        
        for company in companies_km_prop:
            periods = self.ranking_km_proporcional_processor.find_available_periods(company)
            for period in periods:
                # Verificar se Abst_Mot existe
                if not self.check_abst_mot_exists(company, period):
                    logging.warning(f"[SKIP] Ranking_Km_Proporcional {company} - {period}: Abst_Mot nao existe")
                    continue
                
                self.stats['total'] += 1
                if self.process_ranking_km_proporcional(company, period):
                    self.stats['success'] += 1
                else:
                    self.stats['failed'] += 1
        
        # FASE 5: Processar Turnos_Integracao (precisa de Abst_Mot/Detalhado)
        logging.info("\n" + "=" * 80)
        logging.info("FASE 5: Turnos_Integracao")
        logging.info("=" * 80)
        
        companies_turnos = self.turnos_integracao_processor.find_available_companies()
        logging.info(f"Empresas disponiveis para Turnos_Integracao: {len(companies_turnos)}")
        
        for company in companies_turnos:
            periods = self.turnos_integracao_processor.find_available_periods(company)
            for period in periods:
                self.stats['total'] += 1
                if self.process_turnos_integracao(company, period):
                    self.stats['success'] += 1
                else:
                    self.stats['failed'] += 1
        
        # FASE 6: Processar Ranking_Ouro_Mediano (consolidacao)
        logging.info("\n" + "=" * 80)
        logging.info("FASE 6: Ranking_Ouro_Mediano (Consolidacao)")
        logging.info("=" * 80)
        
        if ranking_processed:
            self.stats['total'] += 1
            if self.process_ranking_ouro_mediano():
                self.stats['success'] += 1
            else:
                self.stats['failed'] += 1
        else:
            logging.warning("[SKIP] Ranking_Ouro_Mediano: Nenhum Ranking_Por_Empresa disponivel")
        
        # Resumo final
        elapsed = time.time() - start_time
        self.print_summary(elapsed)
    
    def print_summary(self, elapsed_time: float):
        """Imprime resumo do processamento"""
        logging.info("\n" + "=" * 80)
        logging.info("RESUMO DO PROCESSAMENTO")
        logging.info("=" * 80)
        logging.info(f"Tempo total: {elapsed_time:.2f} segundos")
        logging.info(f"Total de operacoes: {self.stats['total']}")
        logging.info(f"Sucesso: {self.stats['success']}")
        logging.info(f"Falhas: {self.stats['failed']}")
        logging.info(f"Ignorados (ja existem): {self.stats['skipped']}")
        
        if self.stats['total'] > 0:
            rate = (self.stats['success'] / self.stats['total']) * 100
            logging.info(f"Taxa de sucesso: {rate:.1f}%")


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Processamento em Batch de Relatorios')
    parser.add_argument('--entrada', '-e', type=str, default=r"D:\Scripts\Entrada", 
                        help='Diretorio de entrada')
    parser.add_argument('--saida', '-s', type=str, default=r"D:\Scripts\Saida",
                        help='Diretorio de saida')
    parser.add_argument('--versao', '-v', type=str, default="",
                        help='Sufixo de versao (ex: _1.0)')
    
    args = parser.parse_args()
    
    print("=" * 80)
    print("PROCESSAMENTO EM BATCH DE RELATORIOS")
    print("=" * 80)
    
    if not os.path.exists(args.entrada):
        print(f"[ERRO] Diretorio de entrada nao encontrado: {args.entrada}")
        return
    
    if not os.path.exists(args.saida):
        print(f"[ERRO] Diretorio de saida nao encontrado: {args.saida}")
        return
    
    print(f"[ENTRADA] {args.entrada}")
    print(f"[SAIDA] {args.saida}")
    print(f"[VERSAO] {args.versao if args.versao else '(sem sufixo)'}")
    
    processor = BatchProcessor(args.entrada, args.saida, args.versao)
    processor.run_all()
    
    print("\n[LOG] Log completo salvo em: batch_processing.log")


if __name__ == "__main__":
    main()

