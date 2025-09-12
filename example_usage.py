#!/usr/bin/env python3
"""
Example usage of the report generation system.

Demonstrates various ways to use the LaTeX report generator
with sample data and common use cases.
"""

import json
import os
from generate_reports import LaTeXGenerator, filter_inspections


def create_sample_data():
    """Create sample inspection data for demonstration."""
    return [
        {
            "inspection_id": "12345",
            "corporation": "Acme Building Corp",
            "venue": "Downtown Office Complex",
            "building": "Tower A",
            "scheduled_date": "2024-01-15",
            "creation_date": "2024-01-10",
            "completion_date": "2024-01-15",
            "completed_by": "John Smith",
            "overall_comment": "Overall facility in good condition. Minor maintenance items identified.",
            "score_percent": 87,
            "alert_type": "",
            "elements": [
                {
                    "zone": "Lobby",
                    "location": "Main Entrance",
                    "element": "Fire Extinguisher",
                    "score_factor": 1.2,
                    "element_weight_percent": 15,
                    "rating": "Good",
                    "element_score_percent": 90,
                    "comments": "Properly mounted and tagged. Last inspection current.",
                    "attachment": "attachments/12345_img_1.png"
                },
                {
                    "zone": "Stairwell B",
                    "location": "Level 3",
                    "element": "Emergency Lighting",
                    "score_factor": 1.5,
                    "element_weight_percent": 20,
                    "rating": "Fair",
                    "element_score_percent": 75,
                    "comments": "Two bulbs need replacement. Otherwise functional.",
                    "attachment": None
                }
            ]
        },
        {
            "inspection_id": "12346",
            "corporation": "Acme Building Corp",
            "venue": "Warehouse District",
            "building": "Storage Facility C",
            "scheduled_date": "2024-01-16",
            "creation_date": "2024-01-11",
            "completion_date": "2024-01-16",
            "completed_by": "Jane Doe",
            "overall_comment": "Critical safety issues identified. Immediate attention required.",
            "score_percent": 35,
            "alert_type": "Critical",
            "elements": [
                {
                    "zone": "Loading Dock",
                    "location": "Bay 3",
                    "element": "Safety Barriers",
                    "score_factor": 2.0,
                    "element_weight_percent": 40,
                    "rating": "Poor",
                    "element_score_percent": 25,
                    "comments": "Barriers damaged and not secure. Replace immediately.",
                    "attachment": "attachments/12346_img_1.png"
                }
            ]
        }
    ]


def example_single_reports():
    """Example: Generate individual reports for each inspection."""
    print("=== Example: Individual Reports ===")
    
    # Create sample data
    inspections = create_sample_data()
    
    # Save to JSON file for testing
    with open('sample_inspection_data.json', 'w', encoding='utf-8') as f:
        json.dump(inspections, f, indent=2)
    
    # Initialize generator
    generator = LaTeXGenerator(output_dir="example_reports")
    
    # Generate individual reports
    for inspection in inspections:
        tex_file = generator.generate_single_report(inspection)
        print(f"Generated: {tex_file}")
        
        # Optionally compile to PDF (if pdflatex available)
        pdf_file = generator.compile_pdf(tex_file)
        if pdf_file:
            print(f"Compiled: {pdf_file}")


def example_combined_report():
    """Example: Generate combined report with all inspections."""
    print("\\n=== Example: Combined Report ===")
    
    inspections = create_sample_data()
    generator = LaTeXGenerator(output_dir="example_reports")
    
    # Generate combined report
    tex_file = generator.generate_combined_report(inspections)
    print(f"Generated: {tex_file}")
    
    # Compile to PDF
    pdf_file = generator.compile_pdf(tex_file)
    if pdf_file:
        print(f"Compiled: {pdf_file}")


def example_filtering():
    """Example: Filter inspections by various criteria."""
    print("\\n=== Example: Filtering ===")
    
    inspections = create_sample_data()
    
    # Filter by inspector
    john_inspections = filter_inspections(inspections, {'inspector': 'John'})
    print(f"John's inspections: {len(john_inspections)}")
    
    # Filter by alert type
    critical_inspections = filter_inspections(inspections, {'alert_type': 'Critical'})
    print(f"Critical inspections: {len(critical_inspections)}")
    
    # Filter by venue
    warehouse_inspections = filter_inspections(inspections, {'venue': 'Warehouse'})
    print(f"Warehouse inspections: {len(warehouse_inspections)}")
    
    # Generate report for filtered data
    if critical_inspections:
        generator = LaTeXGenerator(output_dir="example_reports")
        tex_file = generator.generate_combined_report(critical_inspections)
        # Rename for clarity
        critical_report = tex_file.parent / "critical_inspections_only.tex"
        tex_file.rename(critical_report)
        print(f"Critical-only report: {critical_report}")


def example_custom_template():
    """Example: Using custom template processing."""
    print("\\n=== Example: Custom Template Processing ===")
    
    # Custom template content (minimal example)
    custom_template = r"""
\documentclass{article}
\usepackage[utf8]{inputenc}
\usepackage{xcolor}

\begin{document}

\title{Custom Inspection Report}
\author{{{COMPLETED_BY}}}
\date{\today}
\maketitle

\section{Inspection {{INSPECTION_ID}}}

\textbf{Location:} {{VENUE}} - {{BUILDING}}

\textbf{Score:} 
{{#if SCORE_PERCENT}}
{{SCORE_PERCENT}}\%
{{/if}}

{{#if ALERT_TYPE}}
\textcolor{red}{\textbf{ALERT: {{ALERT_TYPE}}}}
{{/if}}

\section{Elements}
{{#each ELEMENTS}}
\subsection{{{element}}}
Location: {{location}}\\
Rating: {{rating}}\\
{{#if comments}}
Comments: {{comments}}
{{/if}}

{{/each}}

\end{document}
"""
    
    # Save custom template
    os.makedirs("custom_templates", exist_ok=True)
    with open("custom_templates/simple_template.tex", 'w', encoding='utf-8') as f:
        f.write(custom_template)
    
    # Use custom template
    generator = LaTeXGenerator(template_dir="custom_templates", output_dir="example_reports")
    inspections = create_sample_data()
    
    # Generate using custom template
    for inspection in inspections:
        # Temporarily modify the generate method to use our custom template
        template = generator.load_template("simple_template.tex")
        latex_content = generator.process_template_variables(template, inspection)
        
        # Save output
        inspection_id = inspection.get('inspection_id', 'unknown')
        output_file = generator.output_dir / f"custom_{inspection_id}.tex"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(latex_content)
        
        print(f"Generated custom report: {output_file}")


def example_error_handling():
    """Example: Handling common errors and edge cases."""
    print("\\n=== Example: Error Handling ===")
    
    generator = LaTeXGenerator(output_dir="example_reports")
    
    # Test with missing data
    incomplete_inspection = {
        "inspection_id": "TEST",
        "corporation": None,  # Missing data
        "venue": "",  # Empty string
        "elements": [
            {
                "element": "Test & Special Characters",  # Special chars
                "comments": "Contains $pecial LaTeX character$",
                "score_factor": "invalid_number",  # Invalid number
                "attachment": "nonexistent/image.png"  # Missing image
            }
        ]
    }
    
    try:
        tex_file = generator.generate_single_report(incomplete_inspection)
        print(f"Successfully handled incomplete data: {tex_file}")
    except Exception as e:
        print(f"Error handling incomplete data: {e}")
    
    # Test with empty inspections list
    try:
        empty_report = generator.generate_combined_report([])
        print(f"Generated empty report: {empty_report}")
    except Exception as e:
        print(f"Error with empty list: {e}")


def cleanup_examples():
    """Clean up example files."""
    import shutil
    
    # Remove example directories and files
    for path in ["example_reports", "custom_templates", "sample_inspection_data.json"]:
        if os.path.exists(path):
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
            print(f"Cleaned up: {path}")


def main():
    """Run all examples."""
    print("LaTeX Report Generator Examples")
    print("=" * 50)
    
    try:
        example_single_reports()
        example_combined_report()
        example_filtering()
        example_custom_template()
        example_error_handling()
        
        print("\\n" + "=" * 50)
        print("All examples completed successfully!")
        print("\\nGenerated files are in:")
        print("- example_reports/ - Individual and combined reports")
        print("- custom_templates/ - Custom template example")
        
        # Ask if user wants to clean up
        response = input("\\nClean up example files? (y/N): ").strip().lower()
        if response == 'y':
            cleanup_examples()
            print("Example files cleaned up.")
        else:
            print("Example files preserved for inspection.")
            
    except Exception as e:
        print(f"Example failed: {e}")
        return 1
    
    return 0


if __name__ == '__main__':
    exit(main())