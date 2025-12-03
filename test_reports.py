# -*- coding: utf-8 -*-
"""
Script de Teste para Validacao de Relatorios
=============================================
Este script testa a geracao de todos os relatorios empresa por empresa,
identificando problemas e sugerindo solucoes.

Relatorios validados:
1. Abst_Mot_Por_empresa
2. Ranking_km_Proporcional
3. Ranking_Integracao
4. Ranking_Ouro_Mediano
5. Ranking_Por_Empresa
6. RMC_Destribuida
7. Turnos_Integracao
"""

import os
import sys
import logging
from datetime import datetime
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
import traceback

# Configurar logging para o teste
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('test_reports.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

@dataclass
class ReportDependency:
    """Define as dependencias de cada relatorio"""
    name: str
    input_folders: List[str]
    input_patterns: List[str]
    depends_on_reports: List[str] = field(default_factory=list)
    output_pattern: str = ""

@dataclass
class TestResult:
    """Resultado de um teste"""
    report_type: str
    company: str
    period: str
    success: bool
    error_message: str = ""
    missing_files: List[str] = field(default_factory=list)
    suggested_solutions: List[str] = field(default_factory=list)

class ReportTester:
    """Classe para testar a geracao de relatorios"""
    
    def __init__(self, base_dir: str, output_dir: str):
        self.base_dir = base_dir
        self.output_dir = output_dir
        self.results: List[TestResult] = []
        
        # Definir dependencias de cada relatorio (com acentuacao correta)
        self.report_dependencies = {
            "Abst_Mot_Por_empresa": ReportDependency(
                name="Abst_Mot_Por_empresa",
                input_folders=["Integração_Abast", "Integração_Mot"],
                input_patterns=[
                    "Abastecimento_{company}_{month}_{year}.xlsx",
                    "Motorista_{company}_{month}_{year}.xlsx"
                ],
                depends_on_reports=[],
                output_pattern="Abst_Mot_Por_empresa/{company}/{year}/{month}/Abst_Mot_Por_empresa_{company}_{month}_{year}*.xlsx"
            ),
            "Ranking_Por_Empresa": ReportDependency(
                name="Ranking_Por_Empresa",
                input_folders=["Ranking", "Turnos_128"],
                input_patterns=[
                    "Ranking_{company}_{month}_{year}.xlsx",
                    "Turnos_128_{company}_{month}_{year}.xlsx"
                ],
                depends_on_reports=["Abst_Mot_Por_empresa"],
                output_pattern="Ranking_Por_Empresa/{company}/{year}/{month}/Ranking_Por_Empresa_{company}_{month}_{year}*.xlsx"
            ),
            "Ranking_Integracao": ReportDependency(
                name="Ranking_Integracao",
                input_folders=["Ranking", "Turnos_128"],
                input_patterns=[
                    "Ranking_{company}_{month}_{year}.xlsx",
                    "Turnos_128_{company}_{month}_{year}.xlsx"
                ],
                depends_on_reports=["Abst_Mot_Por_empresa"],
                output_pattern="Ranking_Integração/{company}/{year}/{month}/Ranking_Integração_{company}_{month}_{year}*.xlsx"
            ),
            "Ranking_Ouro_Mediano": ReportDependency(
                name="Ranking_Ouro_Mediano",
                input_folders=[],
                input_patterns=[],
                depends_on_reports=["Ranking_Por_Empresa"],
                output_pattern="Ranking_Ouro_Mediano/Ranking_Ouro_Mediano_*.xlsx"
            ),
            "Ranking_Km_Proporcional": ReportDependency(
                name="Ranking_Km_Proporcional",
                input_folders=[],
                input_patterns=[],
                depends_on_reports=["Abst_Mot_Por_empresa"],
                output_pattern="Rankig_Km_Proporcional/{company}/{year}/{month}/Ranking_Km_Proporcional_{company}_{month}_{year}*.xlsx"
            ),
            "Turnos_Integracao": ReportDependency(
                name="Turnos_Integracao",
                input_folders=[],
                input_patterns=[],
                depends_on_reports=["Abst_Mot_Por_empresa"],
                output_pattern="Turnos Integração/{company}/{year}/{month}/Turnos_Integração_{company}_{month}_{year}*.xlsx"
            ),
            "RMC_Destribuida": ReportDependency(
                name="RMC_Destribuida",
                input_folders=["Resumo_Motorista_Cliente", "Integração_Abast"],
                input_patterns=[
                    "RMC_{company}_{month}_{year}.xlsx",
                    "Abastecimento_{company}_{month}_{year}.xlsx"
                ],
                depends_on_reports=[],
                output_pattern="RMC_Destribuida/{company}/{year}/{month}/RMC_Km_l_Distribuida_{company}_{month}_{year}*.xlsx"
            )
        }
    
    def find_available_companies(self) -> List[str]:
        """Encontra todas as empresas disponiveis"""
        companies = set()
        
        # Verificar pasta de integracao de abastecimento (com acento)
        supply_folder = os.path.join(self.base_dir, "Integração_Abast")
        if os.path.exists(supply_folder):
            for f in os.listdir(supply_folder):
                if f.startswith("Abastecimento_") and f.endswith(".xlsx"):
                    parts = f.replace(".xlsx", "").split("_")
                    if len(parts) >= 2:
                        companies.add(parts[1])
        
        # Verificar pasta de ranking
        ranking_folder = os.path.join(self.base_dir, "Ranking")
        if os.path.exists(ranking_folder):
            for f in os.listdir(ranking_folder):
                if f.startswith("Ranking_") and f.endswith(".xlsx"):
                    parts = f.replace(".xlsx", "").split("_")
                    if len(parts) >= 2 and parts[1] != "Consolidado":
                        companies.add(parts[1])
        
        return sorted(list(companies))
    
    def find_available_periods(self, company: str) -> List[str]:
        """Encontra periodos disponiveis para uma empresa"""
        periods = set()
        
        # Verificar arquivos de abastecimento (com acento)
        supply_folder = os.path.join(self.base_dir, "Integração_Abast")
        if os.path.exists(supply_folder):
            for f in os.listdir(supply_folder):
                if f.startswith(f"Abastecimento_{company}_") and f.endswith(".xlsx"):
                    parts = f.replace(".xlsx", "").split("_")
                    if len(parts) >= 4:
                        month = parts[2]
                        year = parts[3]
                        periods.add(f"{month}_{year}")
        
        return sorted(list(periods))
    
    def check_input_files(self, report_type: str, company: str, period: str) -> Tuple[bool, List[str]]:
        """Verifica se os arquivos de entrada existem"""
        dependency = self.report_dependencies.get(report_type)
        if not dependency:
            return False, [f"Tipo de relatorio desconhecido: {report_type}"]
        
        missing_files = []
        month, year = period.split("_")
        
        for i, folder in enumerate(dependency.input_folders):
            if i < len(dependency.input_patterns):
                pattern = dependency.input_patterns[i]
                filename = pattern.format(company=company, month=month, year=year)
                filepath = os.path.join(self.base_dir, folder, filename)
                
                if not os.path.exists(filepath):
                    missing_files.append(filepath)
        
        return len(missing_files) == 0, missing_files
    
    def check_dependent_reports(self, report_type: str, company: str, period: str) -> Tuple[bool, List[str]]:
        """Verifica se os relatorios dependentes existem"""
        dependency = self.report_dependencies.get(report_type)
        if not dependency:
            return False, []
        
        missing_reports = []
        month, year = period.split("_")
        
        for dep_report in dependency.depends_on_reports:
            dep_dependency = self.report_dependencies.get(dep_report)
            if dep_dependency:
                output_pattern = dep_dependency.output_pattern.format(
                    company=company, 
                    year=year, 
                    month=month.zfill(2)
                )
                output_path = os.path.join(self.output_dir, output_pattern.split("*")[0])
                
                if dep_report == "Abst_Mot_Por_empresa":
                    detalhado_path = os.path.join(
                        self.output_dir, "Abst_Mot_Por_empresa", 
                        company, year, month.zfill(2),
                        f"Detalhado_{company}_{period}.xlsx"
                    )
                    main_path = os.path.join(
                        self.output_dir, "Abst_Mot_Por_empresa",
                        company, year, month.zfill(2),
                        f"Abst_Mot_Por_empresa_{company}_{period}.xlsx"
                    )
                    
                    if not os.path.exists(detalhado_path) and not os.path.exists(main_path):
                        found = False
                        base_dir = os.path.join(self.output_dir, "Abst_Mot_Por_empresa", company, year, month.zfill(2))
                        if os.path.exists(base_dir):
                            for f in os.listdir(base_dir):
                                if f.startswith("Detalhado_") or f.startswith("Abst_Mot_Por_empresa_"):
                                    found = True
                                    break
                        if not found:
                            missing_reports.append(f"{dep_report} (arquivos Abst_Mot_Por_empresa ou Detalhado)")
                else:
                    base_dir = os.path.dirname(output_path)
                    if not os.path.exists(base_dir):
                        missing_reports.append(f"{dep_report} (diretorio nao existe)")
                    else:
                        found = False
                        for f in os.listdir(base_dir):
                            if f.endswith(".xlsx"):
                                found = True
                                break
                        if not found:
                            missing_reports.append(f"{dep_report} (nenhum arquivo encontrado)")
        
        return len(missing_reports) == 0, missing_reports
    
    def check_output_exists(self, report_type: str, company: str, period: str) -> bool:
        """Verifica se o relatorio de saida ja existe"""
        dependency = self.report_dependencies.get(report_type)
        if not dependency:
            return False
        
        month, year = period.split("_")
        
        if report_type == "Ranking_Ouro_Mediano":
            output_dir = os.path.join(self.output_dir, "Ranking_Ouro_Mediano")
            if os.path.exists(output_dir):
                for f in os.listdir(output_dir):
                    if f.startswith("Ranking_Ouro_Mediano_") and f.endswith(".xlsx"):
                        return True
            return False
        
        output_pattern = dependency.output_pattern.format(
            company=company,
            year=year,
            month=month.zfill(2)
        )
        
        base_dir = os.path.join(self.output_dir, os.path.dirname(output_pattern.split("*")[0]))
        
        if os.path.exists(base_dir):
            prefix = os.path.basename(output_pattern.split("*")[0])
            for f in os.listdir(base_dir):
                if f.startswith(prefix) and f.endswith(".xlsx"):
                    return True
        
        return False
    
    def generate_solutions(self, report_type: str, missing_files: List[str], missing_reports: List[str]) -> List[str]:
        """Gera solucoes para os problemas encontrados"""
        solutions = []
        
        if missing_files:
            solutions.append("[ARQUIVOS] ARQUIVOS DE ENTRADA FALTANDO:")
            for f in missing_files:
                solutions.append(f"   [X] {f}")
            solutions.append("")
            solutions.append("[!] SOLUCAO: Verifique se os arquivos de entrada estao presentes nas pastas corretas:")
            
            if "Integração_Abast" in str(missing_files) or "Integracao_Abast" in str(missing_files):
                solutions.append("   - Pasta Integração_Abast: deve conter arquivos 'Abastecimento_EMPRESA_MES_ANO.xlsx'")
            if "Integração_Mot" in str(missing_files) or "Integracao_Mot" in str(missing_files):
                solutions.append("   - Pasta Integração_Mot: deve conter arquivos 'Motorista_EMPRESA_MES_ANO.xlsx'")
            if "Ranking" in str(missing_files):
                solutions.append("   - Pasta Ranking: deve conter arquivos 'Ranking_EMPRESA_MES_ANO.xlsx'")
            if "Turnos_128" in str(missing_files):
                solutions.append("   - Pasta Turnos_128: deve conter arquivos 'Turnos_128_EMPRESA_MES_ANO.xlsx'")
            if "Resumo_Motorista_Cliente" in str(missing_files):
                solutions.append("   - Pasta Resumo_Motorista_Cliente: deve conter arquivos 'RMC_EMPRESA_MES_ANO.xlsx'")
        
        if missing_reports:
            solutions.append("")
            solutions.append("[DEPENDENCIAS] RELATORIOS DEPENDENTES FALTANDO:")
            for r in missing_reports:
                solutions.append(f"   [X] {r}")
            solutions.append("")
            solutions.append("[!] SOLUCAO: Gere os relatorios dependentes primeiro na seguinte ordem:")
            
            order = {
                "Abst_Mot_Por_empresa": 1,
                "Ranking_Por_Empresa": 2,
                "Ranking_Integracao": 3,
                "Ranking_Km_Proporcional": 4,
                "Turnos_Integracao": 5,
                "Ranking_Ouro_Mediano": 6,
                "RMC_Destribuida": 7
            }
            
            missing_ordered = sorted(missing_reports, key=lambda x: order.get(x.split()[0], 99))
            for i, r in enumerate(missing_ordered, 1):
                solutions.append(f"   {i}. {r}")
        
        return solutions
    
    def test_report(self, report_type: str, company: str, period: str) -> TestResult:
        """Testa a geracao de um relatorio especifico"""
        result = TestResult(
            report_type=report_type,
            company=company,
            period=period,
            success=False
        )
        
        try:
            inputs_ok, missing_files = self.check_input_files(report_type, company, period)
            result.missing_files.extend(missing_files)
            
            deps_ok, missing_reports = self.check_dependent_reports(report_type, company, period)
            
            if missing_files or missing_reports:
                result.suggested_solutions = self.generate_solutions(report_type, missing_files, missing_reports)
            
            output_exists = self.check_output_exists(report_type, company, period)
            
            if output_exists:
                result.success = True
                result.error_message = "Relatorio ja existe"
            elif not inputs_ok:
                result.error_message = f"Arquivos de entrada faltando: {len(missing_files)}"
            elif not deps_ok:
                result.error_message = f"Relatorios dependentes faltando: {', '.join(missing_reports)}"
            else:
                result.success = True
                result.error_message = "Pronto para gerar (arquivos de entrada OK)"
        
        except Exception as e:
            result.error_message = f"Erro durante validacao: {str(e)}"
            result.suggested_solutions = [f"[!] Erro: {traceback.format_exc()}"]
        
        return result
    
    def run_all_tests(self, companies: Optional[List[str]] = None, periods: Optional[List[str]] = None):
        """Executa todos os testes"""
        if companies is None:
            companies = self.find_available_companies()
        
        if not companies:
            logging.warning("Nenhuma empresa encontrada para teste")
            return
        
        logging.info("=" * 80)
        logging.info("INICIANDO VALIDACAO DE RELATORIOS")
        logging.info(f"Empresas: {len(companies)}")
        logging.info("=" * 80)
        
        report_types = [
            "Abst_Mot_Por_empresa",
            "Ranking_Por_Empresa", 
            "Ranking_Integracao",
            "Ranking_Ouro_Mediano",
            "Ranking_Km_Proporcional",
            "Turnos_Integracao",
            "RMC_Destribuida"
        ]
        
        for company in companies:
            logging.info("")
            logging.info("=" * 80)
            logging.info(f"EMPRESA: {company}")
            logging.info("=" * 80)
            
            company_periods = periods if periods else self.find_available_periods(company)
            
            if not company_periods:
                logging.warning(f"Nenhum periodo encontrado para {company}")
                continue
            
            for period in company_periods:
                logging.info("")
                logging.info(f"[PERIODO] {period}")
                logging.info("-" * 40)
                
                for report_type in report_types:
                    result = self.test_report(report_type, company, period)
                    self.results.append(result)
                    
                    status = "[OK]" if result.success else "[ERRO]"
                    logging.info(f"  {status} {report_type}: {result.error_message}")
                    
                    if not result.success and result.suggested_solutions:
                        for solution in result.suggested_solutions:
                            logging.info(f"      {solution}")
    
    def generate_summary(self) -> str:
        """Gera um resumo dos testes"""
        summary = []
        summary.append("")
        summary.append("=" * 80)
        summary.append("RESUMO DOS TESTES")
        summary.append("=" * 80)
        
        total = len(self.results)
        success = len([r for r in self.results if r.success])
        failed = total - success
        
        summary.append(f"\nTotal de testes: {total}")
        summary.append(f"[OK] Sucesso: {success}")
        summary.append(f"[ERRO] Falha: {failed}")
        summary.append(f"Taxa de sucesso: {(success/total*100):.1f}%" if total > 0 else "N/A")
        
        summary.append("")
        summary.append("-" * 40)
        summary.append("RESUMO POR TIPO DE RELATORIO")
        summary.append("-" * 40)
        
        report_types = set(r.report_type for r in self.results)
        for report_type in sorted(report_types):
            type_results = [r for r in self.results if r.report_type == report_type]
            type_success = len([r for r in type_results if r.success])
            type_total = len(type_results)
            status = "[OK]" if type_success == type_total else "[ERRO]"
            summary.append(f"  {status} {report_type}: {type_success}/{type_total}")
        
        summary.append("")
        summary.append("-" * 40)
        summary.append("RESUMO POR EMPRESA")
        summary.append("-" * 40)
        
        companies = set(r.company for r in self.results)
        for company in sorted(companies):
            company_results = [r for r in self.results if r.company == company]
            company_success = len([r for r in company_results if r.success])
            company_total = len(company_results)
            status = "[OK]" if company_success == company_total else "[PARCIAL]" if company_success > 0 else "[ERRO]"
            summary.append(f"  {status} {company}: {company_success}/{company_total}")
        
        summary.append("")
        summary.append("-" * 40)
        summary.append("PROBLEMAS MAIS COMUNS")
        summary.append("-" * 40)
        
        error_counts = {}
        for r in self.results:
            if not r.success and r.error_message:
                error_key = r.error_message.split(":")[0] if ":" in r.error_message else r.error_message
                error_counts[error_key] = error_counts.get(error_key, 0) + 1
        
        for error, count in sorted(error_counts.items(), key=lambda x: -x[1]):
            summary.append(f"  [X] {error}: {count} ocorrencia(s)")
        
        return "\n".join(summary)


def main():
    """Funcao principal"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Validacao de Relatorios')
    parser.add_argument('--entrada', '-e', type=str, help='Diretorio de entrada')
    parser.add_argument('--saida', '-s', type=str, help='Diretorio de saida')
    parser.add_argument('--empresa', '-c', type=str, help='Empresas (separadas por virgula)')
    parser.add_argument('--periodo', '-p', type=str, help='Periodos (separados por virgula)')
    parser.add_argument('--auto', '-a', action='store_true', help='Executar automaticamente sem interacao')
    
    args = parser.parse_args()
    
    print("=" * 80)
    print("SCRIPT DE VALIDACAO DE RELATORIOS")
    print("=" * 80)
    
    possible_base_dirs = [
        r"D:\Scripts\Entrada",
        r"D:\Scripts\Integracao_Abast",
        os.getcwd()
    ]
    
    possible_output_dirs = [
        r"D:\Scripts\Saida",
        os.getcwd()
    ]
    
    base_dir = args.entrada
    output_dir = args.saida
    
    if not base_dir:
        for d in possible_base_dirs:
            if os.path.exists(d):
                base_dir = d
                break
    
    if not output_dir:
        for d in possible_output_dirs:
            if os.path.exists(d):
                output_dir = d
                break
    
    if not base_dir or not output_dir:
        print("\n[!] CONFIGURACAO NECESSARIA:")
        print("Por favor, use os parametros --entrada e --saida")
        print("  Exemplo: python test_reports.py --entrada D:\\Scripts\\Entrada --saida D:\\Scripts\\Saida --auto")
        return
    
    if not os.path.exists(base_dir):
        print(f"[ERRO] Diretorio de entrada nao encontrado: {base_dir}")
        return
    
    if not os.path.exists(output_dir):
        print(f"[ERRO] Diretorio de saida nao encontrado: {output_dir}")
        return
    
    print(f"\n[ENTRADA] Diretorio de entrada: {base_dir}")
    print(f"[SAIDA] Diretorio de saida: {output_dir}")
    
    tester = ReportTester(base_dir, output_dir)
    
    companies = None
    periods = None
    
    if args.empresa:
        companies = [c.strip() for c in args.empresa.split(",")]
    if args.periodo:
        periods = [p.strip() for p in args.periodo.split(",")]
    
    print("\n[...] Iniciando validacao...")
    tester.run_all_tests(companies, periods)
    
    summary = tester.generate_summary()
    print(summary)
    logging.info(summary)
    
    print(f"\n[LOG] Log completo salvo em: test_reports.log")


if __name__ == "__main__":
    main()
