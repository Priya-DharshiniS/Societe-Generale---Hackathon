from pyparsing import Word, alphas, alphanums, nums, Suppress, Group, ZeroOrMore, Keyword, LineEnd, Optional, Combine

def parse_vba_code(vba_code):
    identifier = Word(alphas, alphanums + '_')
    integer = Word(nums)
    string = Word(alphas + alphanums + '_')
    line_end = Suppress(LineEnd())
    
    # Keywords
    sub_keyword = Keyword("Sub")
    end_sub_keyword = Keyword("End Sub")
    with_keyword = Keyword("With")
    end_with_keyword = Keyword("End With")
    selection_keyword = Keyword("Selection")
    range_keyword = Keyword("Range")
    dim_keyword = Keyword("Dim")
    as_keyword = Keyword("As")
    if_keyword = Keyword("If")
    then_keyword = Keyword("Then")
    else_keyword = Keyword("Else")
    endif_keyword = Keyword("End If")
    for_keyword = Keyword("For")
    to_keyword = Keyword("To")
    next_keyword = Keyword("Next")
    
    # Simple grammar for demonstration
    sub_declaration = Group(sub_keyword + identifier + Suppress('(') + Suppress(')') + line_end)
    end_sub = Group(end_sub_keyword + line_end)
    with_statement = Group(with_keyword + selection_keyword + Suppress('.') + identifier + line_end + end_with_keyword)
    range_statement = Group(identifier + Suppress('.') + range_keyword + Suppress('(') + Suppress('"') + identifier + Suppress('"') + Suppress(')') + Suppress('.') + identifier + Suppress('(') + integer + Suppress(')') + line_end)
    assignment_statement = Group(identifier + Suppress('=') + (integer | string) + line_end)
    dim_statement = Group(dim_keyword + identifier + Optional(as_keyword + identifier) + line_end)
    if_statement = Group(if_keyword + identifier + Suppress('=') + integer + then_keyword + line_end)
    for_statement = Group(for_keyword + identifier + Suppress('=') + integer + to_keyword + integer + line_end)
    next_statement = Group(next_keyword + identifier + line_end)
    
    vba_grammar = ZeroOrMore(sub_declaration | end_sub | with_statement | range_statement | assignment_statement | dim_statement | if_statement | for_statement | next_statement)
    
    ast = vba_grammar.parseString(vba_code)
    print(f"Parsed AST: {ast}")  # Debugging statement
    return ast
