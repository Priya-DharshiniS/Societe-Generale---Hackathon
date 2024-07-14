from vba_parser import parse_vba_code
from documentation_generator import generate_documentation

def analyze_vba_code(vba_code):
    analysis_results = {}
    for module_name, code in vba_code.items():
        ast = parse_vba_code(code)
        analysis = {
            "logic": analyze_logic(ast),
            "data_flow": analyze_data_flow(ast),
            "process_flow": analyze_process_flow(ast)
        }
        analysis_results[module_name] = analysis
    return analysis_results

def analyze_logic(ast):
    logic_components = []
    for item in ast:
        if item[0] == 'If' or item[0] == 'Then' or item[0] == 'Else' or item[0] == 'End If':
            logic_components.append(f"Conditional detected: {' '.join(item)}")
        if item[0] == 'For' or item[0] == 'Next':
            logic_components.append(f"Loop detected: {' '.join(item)}")
    return "\n".join(logic_components) if logic_components else "No complex logic (loops, conditionals, etc.) found."

def analyze_data_flow(ast):
    variables = set()
    for item in ast:
        if item[0] == 'Dim':
            var_name = item[1]
            variables.add(var_name)
    return f"Variables: {', '.join(variables)}" if variables else "No variables found."

def analyze_process_flow(ast):
    process_flow = []
    for item in ast:
        if item[0] == 'Sub' or item[0] == 'End Sub':
            process_flow.append(f"Subroutine detected: {' '.join(item)}")
        if item[0] == 'With' or item[0] == 'End With':
            process_flow.append(f"With block detected: {' '.join(item)}")
        if item[0] == 'Range':
            process_flow.append(f"Range operation detected: {' '.join(item)}")
    return "\n".join(process_flow) if process_flow else "No process steps found."

# Example usage (for testing purposes)
if __name__ == "__main__":
    example_code = {
        "Module1": """
        Sub Module1Example()
            Dim i As Integer
            Dim result As Integer
            
            For i = 1 To 10
                If i Mod 2 = 0 Then
                    result = i * 2
                Else
                    result = i * 3
                End If
            Next i
            
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 192, 0)  ' Example: Orange color
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
            ActiveCell.Offset(0, 2).Range("A1").Select
        End Sub""",
        "Module2": """
        Sub Module2Example()
            Dim x As Integer
            Dim y As Integer
            
            x = 10
            y = 20
            
            If x > y Then
                MsgBox "x is greater than y"
            Else
                MsgBox "y is greater than or equal to x"
            End If
        End Sub"""
    }
    analysis = analyze_vba_code(example_code)
    documentation = generate_documentation(analysis)
    print(documentation)
