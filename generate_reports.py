#!/usr/bin/env python3
"""
JSON to LaTeX Report Generator

Converts inspection_summary.json to LaTeX reports using templates.
Supports both individual inspection reports and combined reports.

Usage:
    python generate_reports.py --single                    # Generate individual .tex files
    python generate_reports.py --combined                  # Generate combined report
    python generate_reports.py --compile                   # Generate and compile PDFs
    python generate_reports.py --filter inspector="John"   # Filter by inspector
"""

import json
import os
import re
import subprocess
import argparse
from pathlib import Path
from typing import Dict, List, Any, Optional
from datetime import datetime
import shutil


class LaTeXGenerator:
    """Generates LaTeX reports from JSON inspection data."""
    
    def __init__(self, template_dir: str = "templates", output_dir: str = "reports"):
        self.template_dir = Path(template_dir)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
    def generate_flat_table_report(self, table_data: List[Dict[str, Any]], 
                                   template_name: str = "flat_table_template.tex") -> Path:
        """Generate a flat table report from inspection table data."""
        template = self.load_template(template_name)
        
        # Process the flat table data for LaTeX
        latex_content = self.process_flat_table_template(template, table_data)
        
        # Create output filename
        output_file = self.output_dir / "flat_table_report.tex"
        
        # Write LaTeX file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(latex_content)
        
        print(f"Generated flat table report: {output_file}")
        return output_file
    
    def process_flat_table_template(self, template: str, table_data: List[Dict[str, Any]]) -> str:
        """Process flat table template with table data."""
        # Process {{#each INSPECTION_DATA}} blocks
        pattern = r'\{\{#each\s+INSPECTION_DATA\}\}(.*?)\{\{/each\}\}'
        
        def replace_data_loop(match):
            loop_template = match.group(1)
            result = ""
            
            for row in table_data:
                row_content = loop_template
                
                # Replace all row variables and handle conditionals
                for key, value in row.items():
                    if value is None:
                        value = ""
                    
                    # Handle special formatting
                    if isinstance(value, str) and value:
                        # Don't escape LaTeX characters in file paths
                        if key == 'attachment' and value:
                            formatted_value = value  # Keep file paths as-is
                        else:
                            formatted_value = self.sanitize_latex(value)
                    else:
                        formatted_value = str(value) if value else ""
                    
                    # Replace simple variables
                    row_content = row_content.replace(f"{{{{{key}}}}}", formatted_value)
                    
                    # Handle conditionals for this key
                    if_pattern = r'\{\{#if\s+' + key + r'\}\}(.*?)\{\{/if\}\}'
                    def replace_if(match):
                        if_content = match.group(1)
                        return if_content if formatted_value and formatted_value.strip() else ""
                    row_content = re.sub(if_pattern, replace_if, row_content, flags=re.DOTALL)
                
                result += row_content
            
            return result
        
        return re.sub(pattern, replace_data_loop, template, flags=re.DOTALL)
        
    def load_template(self, template_name: str) -> str:
        """Load LaTeX template from file."""
        template_path = self.template_dir / template_name
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        
        with open(template_path, 'r', encoding='utf-8') as f:
            return f.read()
    
    def sanitize_latex(self, text: str) -> str:
        """Escape special LaTeX characters in text."""
        if not isinstance(text, str):
            text = str(text) if text is not None else ""
        
        # LaTeX special characters that need escaping
        # Order matters - do backslashes first to avoid double-escaping
        latex_chars = [
            ('\\', r'\textbackslash{}'),
            ('&', r'\&'),
            ('%', r'\%'),
            ('$', r'\$'),
            ('#', r'\#'),
            ('^', r'\textasciicircum{}'),
            ('_', r'\_'),
            ('{', r'\{'),
            ('}', r'\}'),
            ('~', r'\textasciitilde{}'),
        ]
        
        for char, escaped in latex_chars:
            text = text.replace(char, escaped)
        
        return text
    
    def format_date(self, date_str: str) -> str:
        """Format date string for LaTeX output."""
        if not date_str or date_str == "None":
            return "Not specified"
        return self.sanitize_latex(str(date_str))
    
    def format_score(self, score: Any) -> str:
        """Format score for LaTeX with appropriate coloring."""
        if score is None:
            return "N/A"
        
        try:
            score_num = float(score)
            return f"{score_num:.0f}"
        except (ValueError, TypeError):
            return "N/A"
    
    def process_template_variables(self, template: str, inspection: Dict[str, Any]) -> str:
        """Replace template variables with inspection data."""
        # Map JSON keys to template variables
        variables = {
            'INSPECTION_ID': self.sanitize_latex(inspection.get('inspection_id', '')),
            'CORPORATION': self.sanitize_latex(inspection.get('corporation', '')),
            'VENUE': self.sanitize_latex(inspection.get('venue', '')),
            'BUILDING': self.sanitize_latex(inspection.get('building', '')),
            'SCHEDULED_DATE': self.format_date(inspection.get('scheduled_date', '')),
            'CREATION_DATE': self.format_date(inspection.get('creation_date', '')),
            'COMPLETION_DATE': self.format_date(inspection.get('completion_date', '')),
            'COMPLETED_BY': self.sanitize_latex(inspection.get('completed_by', '')),
            'OVERALL_COMMENT': self.sanitize_latex(inspection.get('overall_comment', '')),
            'SCORE_PERCENT': self.format_score(inspection.get('score_percent')),
            'ALERT_TYPE': self.sanitize_latex(inspection.get('alert_type', '')),
        }
        
        # Replace simple variables
        latex_content = template
        for var, value in variables.items():
            latex_content = latex_content.replace(f"{{{{{var}}}}}", value or "")
        
        # Process conditional blocks
        latex_content = self.process_conditionals(latex_content, variables)
        
        # Process elements loop
        latex_content = self.process_elements_loop(latex_content, inspection.get('elements', []))
        
        return latex_content
    
    def process_conditionals(self, content: str, variables: Dict[str, str]) -> str:
        """Process {{#if VARIABLE}} blocks."""
        # Simple regex for conditional blocks
        pattern = r'\{\{#if\s+([A-Z_]+)\}\}(.*?)\{\{/if\}\}'
        
        def replace_conditional(match):
            var_name = match.group(1)
            block_content = match.group(2)
            
            if var_name in variables and variables[var_name] and variables[var_name] != "N/A":
                return block_content
            return ""
        
        return re.sub(pattern, replace_conditional, content, flags=re.DOTALL)
    
    def process_elements_loop(self, content: str, elements: List[Dict[str, Any]]) -> str:
        """Process {{#each ELEMENTS}} blocks."""
        # Find the elements loop
        pattern = r'\{\{#each\s+ELEMENTS\}\}(.*?)\{\{/each\}\}'
        
        def replace_elements_loop(match):
            loop_template = match.group(1)
            result = ""
            
            for i, element in enumerate(elements):
                element_content = loop_template
                
                # Replace element variables
                element_vars = {
                    'zone': self.sanitize_latex(element.get('zone', '')),
                    'location': self.sanitize_latex(element.get('location', '')),
                    'element': self.sanitize_latex(element.get('element', '')),
                    'score_factor': str(element.get('score_factor', '')) if element.get('score_factor') is not None else '',
                    'element_weight_percent': str(element.get('element_weight_percent', '')) if element.get('element_weight_percent') is not None else '',
                    'rating': self.sanitize_latex(element.get('rating', '')),
                    'element_score_percent': self.format_score(element.get('element_score_percent')),
                    'comments': self.sanitize_latex(element.get('comments', '')),
                    'attachment': element.get('attachment', ''),
                    '@index_plus_1': str(i + 1),
                }
                
                # Replace variables
                for var, value in element_vars.items():
                    element_content = element_content.replace(f"{{{{{var}}}}}", value or "")
                
                # Process element conditionals
                element_content = self.process_element_conditionals(element_content, element_vars)
                
                result += element_content
            
            return result
        
        return re.sub(pattern, replace_elements_loop, content, flags=re.DOTALL)
    
    def process_element_conditionals(self, content: str, element_vars: Dict[str, str]) -> str:
        """Process conditionals within element loops."""
        pattern = r'\{\{#if\s+([a-zA-Z_@]+)\}\}(.*?)\{\{/if\}\}'
        
        def replace_conditional(match):
            var_name = match.group(1)
            block_content = match.group(2)
            
            if var_name in element_vars and element_vars[var_name] and element_vars[var_name] != "N/A":
                return block_content
            return ""
        
        return re.sub(pattern, replace_conditional, content, flags=re.DOTALL)
    
    def generate_single_report(self, inspection: Dict[str, Any], template_name: str = "inspection_template.tex") -> Path:
        """Generate a single inspection report."""
        template = self.load_template(template_name)
        latex_content = self.process_template_variables(template, inspection)
        
        # Create output filename
        inspection_id = inspection.get('inspection_id', 'unknown')
        safe_id = re.sub(r'[^\w\-_.]', '_', str(inspection_id))
        output_file = self.output_dir / f"inspection_{safe_id}.tex"
        
        # Write LaTeX file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(latex_content)
        
        print(f"Generated: {output_file}")
        return output_file
    
    def generate_combined_report(self, inspections: List[Dict[str, Any]], 
                                template_name: str = "combined_template.tex") -> Path:
        """Generate a combined report with all inspections."""
        # For now, use the single template in a loop structure
        # In production, you'd have a separate combined template
        
        output_file = self.output_dir / "combined_inspection_report.tex"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            # Write document header
            f.write(r"""
\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{geometry}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{longtable}
\usepackage{xcolor}
\usepackage{fancyhdr}
\usepackage{hyperref}

\geometry{margin=2cm, headheight=1.5cm}
\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\textbf{Combined Audit Report}}
\fancyhead[R]{\thepage}
\fancyfoot[C]{\footnotesize Generated: """ + datetime.now().strftime("%Y-%m-%d") + r"""}

\definecolor{headerblue}{RGB}{41, 98, 167}
\definecolor{darkgray}{RGB}{108, 117, 125}

\newcommand{\sectionheader}[1]{%
    \vspace{12pt}%
    {\Large\bfseries\color{headerblue}#1}%
    \vspace{6pt}%
    \hrule height 1pt%
    \vspace{12pt}%
}

\begin{document}

\title{\Huge\bfseries Combined Inspection Report}
\date{\today}
\maketitle

\tableofcontents
\newpage

""")
            
            # Generate each inspection as a section
            for i, inspection in enumerate(inspections):
                inspection_id = inspection.get('inspection_id', f'Unknown_{i+1}')
                f.write(f"\\section{{Inspection {self.sanitize_latex(str(inspection_id))}}}\n\n")
                
                # Use simplified template for combined report
                template_content = self.create_simplified_template()
                latex_content = self.process_template_variables(template_content, inspection)
                f.write(latex_content)
                f.write("\n\\newpage\n\n")
            
            f.write("\\end{document}")
        
        print(f"Generated combined report: {output_file}")
        return output_file
    
    def create_simplified_template(self) -> str:
        """Create a simplified template for combined reports."""
        return r"""
\begin{tabular}{@{}p{4cm}p{10cm}@{}}
    \textbf{Corporation:} & {{CORPORATION}} \\[4pt]
    \textbf{Venue:} & {{VENUE}} \\[4pt]
    \textbf{Building:} & {{BUILDING}} \\[4pt]
    \textbf{Inspector:} & {{COMPLETED_BY}} \\[4pt]
    \textbf{Date:} & {{COMPLETION_DATE}} \\[4pt]
    {{#if SCORE_PERCENT}}
    \textbf{Score:} & {{SCORE_PERCENT}}\% \\[4pt]
    {{/if}}
    {{#if ALERT_TYPE}}
    \textbf{Alert:} & \textcolor{red}{\textbf{{{ALERT_TYPE}}}} \\[4pt]
    {{/if}}
\end{tabular}

{{#if OVERALL_COMMENT}}
\vspace{8pt}
\textbf{Comments:} {{OVERALL_COMMENT}}
{{/if}}

\vspace{12pt}
\textbf{Elements:}

\begin{itemize}
{{#each ELEMENTS}}
\item \textbf{{{element}}} 
    {{#if location}}({{location}}){{/if}}
    {{#if element_score_percent}}-- Score: {{element_score_percent}}\%{{/if}}
    {{#if comments}}\\\\{\footnotesize\textit{{{comments}}}}{{/if}}
{{/each}}
\end{itemize}
"""
    
    def compile_pdf(self, tex_file: Path) -> Optional[Path]:
        """Compile LaTeX file to PDF using pdflatex."""
        
        # Try Python pdflatex package first
        try:
            import pdflatex
            print(f"Using Python pdflatex package to compile {tex_file}")
            
            pdftex = pdflatex.PDFLaTeX.from_texfile(str(tex_file))
            pdf, log, completed_process = pdftex.create_pdf(keep_pdf_file=True, keep_log_file=False)
            
            pdf_file = tex_file.with_suffix('.pdf')
            if pdf_file.exists():
                print(f"Compiled PDF: {pdf_file}")
                return pdf_file
            else:
                print("PDF compilation succeeded but file not found")
                return None
                
        except ImportError:
            print("Python pdflatex package not available, trying system pdflatex")
        except Exception as e:
            print(f"Python pdflatex compilation failed: {e}")
        
        # Fallback to system pdflatex
        pdflatex_cmd = shutil.which('pdflatex')
        
        # Try common MiKTeX locations on Windows
        if not pdflatex_cmd and os.name == 'nt':
            import getpass
            username = getpass.getuser()
            miktex_paths = [
                rf"C:\Users\{username}\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdflatex.exe",
                r"C:\Program Files\MiKTeX\miktex\bin\x64\pdflatex.exe",
                r"C:\Program Files (x86)\MiKTeX\miktex\bin\pdflatex.exe",
            ]
            for path in miktex_paths:
                if os.path.exists(path):
                    pdflatex_cmd = path
                    break
        
        if not pdflatex_cmd:
            print("Warning: Neither Python pdflatex package nor system pdflatex found.")
            print("Install a LaTeX distribution (MiKTeX/TeX Live) for PDF compilation.")
            return None
        
        try:
            print(f"Using pdflatex: {pdflatex_cmd}")
            # Run pdflatex twice for proper references
            for _ in range(2):
                result = subprocess.run([
                    pdflatex_cmd, 
                    '-interaction=nonstopmode',
                    '-output-directory', str(tex_file.parent),
                    tex_file.name  # Use just filename, not full path
                ], capture_output=True, text=True, cwd=tex_file.parent)
                
                if result.returncode != 0:
                    print(f"LaTeX compilation error for {tex_file}:")
                    print(result.stdout)
                    print(result.stderr)
                    return None
            
            pdf_file = tex_file.with_suffix('.pdf')
            if pdf_file.exists():
                print(f"Compiled PDF: {pdf_file}")
                return pdf_file
            
        except Exception as e:
            print(f"Error compiling {tex_file}: {e}")
        
        return None


def filter_inspections(inspections: List[Dict[str, Any]], filters: Dict[str, str]) -> List[Dict[str, Any]]:
    """Filter inspections based on criteria."""
    filtered = inspections
    
    for key, value in filters.items():
        if key == 'inspector':
            filtered = [insp for insp in filtered 
                       if value.lower() in insp.get('completed_by', '').lower()]
        elif key == 'venue':
            filtered = [insp for insp in filtered 
                       if value.lower() in insp.get('venue', '').lower()]
        elif key == 'alert_type':
            filtered = [insp for insp in filtered 
                       if insp.get('alert_type', '').lower() == value.lower()]
    
    return filtered


def main():
    parser = argparse.ArgumentParser(description='Generate LaTeX reports from inspection JSON')
    parser.add_argument('--input', '-i', default='inspection_summary.json',
                       help='Input JSON file (default: inspection_summary.json)')
    parser.add_argument('--output-dir', '-o', default='reports',
                       help='Output directory (default: reports)')
    parser.add_argument('--single', action='store_true',
                       help='Generate individual .tex files per inspection')
    parser.add_argument('--combined', action='store_true',
                       help='Generate combined report')
    parser.add_argument('--compile', action='store_true',
                       help='Compile LaTeX to PDF')
    parser.add_argument('--filter', action='append', nargs=1, metavar='KEY=VALUE',
                       help='Filter inspections (e.g., --filter inspector=John)')
    
    args = parser.parse_args()
    
    # Default to single reports if no mode specified
    if not args.single and not args.combined:
        args.single = True
    
    # Load inspection data
    try:
        with open(args.input, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"Error: Input file '{args.input}' not found.")
        return 1
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in '{args.input}': {e}")
        return 1
    
    if not data:
        print("No data found in input file.")
        return 1
    
    # Detect data format - flat table vs nested inspections
    is_flat_table = (isinstance(data, list) and len(data) > 0 and 
                     'inspection_number' in data[0] and 'element' in data[0])
    
    if is_flat_table:
        print(f"Detected flat table format with {len(data)} rows")
        inspections = data  # Use data directly for flat table
    else:
        print(f"Detected nested inspection format with {len(data)} inspections")
        inspections = data
    
    # Apply filters
    filters = {}
    if args.filter:
        for filter_arg in args.filter:
            key, value = filter_arg[0].split('=', 1)
            filters[key] = value
    
    if filters:
        inspections = filter_inspections(inspections, filters)
        print(f"Applied filters: {filters}")
        print(f"Found {len(inspections)} matching inspections")
    
    # Initialize generator
    generator = LaTeXGenerator(output_dir=args.output_dir)
    generated_files = []
    
    # Generate reports based on data format
    if is_flat_table:
        print("Generating flat table report...")
        tex_file = generator.generate_flat_table_report(inspections)
        generated_files.append(tex_file)
    else:
        # Original nested format logic
        if args.single:
            print(f"Generating individual reports for {len(inspections)} inspections...")
            for inspection in inspections:
                tex_file = generator.generate_single_report(inspection)
                generated_files.append(tex_file)
        
        if args.combined:
            print("Generating combined report...")
            tex_file = generator.generate_combined_report(inspections)
            generated_files.append(tex_file)
    
    # Compile PDFs if requested
    if args.compile:
        print("Compiling LaTeX files to PDF...")
        for tex_file in generated_files:
            generator.compile_pdf(tex_file)
    
    print(f"\\nGenerated {len(generated_files)} LaTeX files in '{args.output_dir}' directory")
    print("\\nNext steps:")
    print("1. Review the generated .tex files")
    print("2. Compile with: pdflatex filename.tex")
    print("3. Or run with --compile flag to auto-compile")
    
    return 0


if __name__ == '__main__':
    exit(main())