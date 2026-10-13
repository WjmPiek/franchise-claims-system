"""Bounded header discovery for monthly policy transaction workbooks."""
import re

ALIASES = {
    'franchise': ('Franchise', 'Branch', 'Branch Name', 'Participating Group', 'Franchise Name', 'Office'),
    'relation': ('Relation', 'Member Relation', 'Rel', 'Member Type'),
    'risk': ('AUL Risk', 'Risk', 'Risk Value', 'Risk Premium', 'Risk Amount'),
    'retail': ('Retail', 'Retail Premium', 'Retail Amount'),
    'mpia': ('MPIA',),
}

def key(value):
    return re.sub(r'[^a-z0-9]+', ' ', str(value or '').strip().lower()).strip()

def policy_columns(header):
    lookup = {key(value): index for index, value in enumerate(header) if key(value)}
    return {field: next((lookup[key(alias)] for alias in aliases if key(alias) in lookup), None)
            for field, aliases in ALIASES.items()}

def find_policy_sheet(book):
    """Read at most 25 header rows per sheet; never guess money column positions."""
    for name in book.sheetnames:
        sheet = book[name]
        first = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        columns = policy_columns(first)
        if all(index is not None for index in columns.values()):
            return sheet, 1, first, columns
        for row_number, header in enumerate(sheet.iter_rows(min_row=2, max_row=25, values_only=True), 2):
            columns = policy_columns(header)
            if all(index is not None for index in columns.values()):
                return sheet, row_number, header, columns
    raise ValueError('No policy transaction header found in the first 25 rows of any worksheet. '
                     'Required columns: Franchise (or Branch), Relation, AUL Risk (or Risk Premium), '
                     'Retail (or Retail Premium), MPIA.')
