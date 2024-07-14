def generate_documentation(analysis_results):
    documentation = []
    for module_name, analysis in analysis_results.items():
        documentation.append(f"Module: {module_name}\n")
        documentation.append("Logic Analysis:\n")
        documentation.append(f"{analysis.get('logic', 'No logic analysis available.')}\n")
        documentation.append("Data Flow Analysis:\n")
        documentation.append(f"{analysis.get('data_flow', 'No data flow analysis available.')}\n")
        documentation.append("Process Flow Analysis:\n")
        documentation.append(f"{analysis.get('process_flow', 'No process flow analysis available.')}\n")
        documentation.append("\n" + "="*40 + "\n")
    return "\n".join(documentation)
