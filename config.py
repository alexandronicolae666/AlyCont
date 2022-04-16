# value != None => default value for the field
# value == None => value has to be extracted from the document
# position_original_document => after formating the values from original document ( and merged values )
HEADERS = {
    'Numar linie': {
        'value': '1',
        'required': True,
        'position_original_document': None
    },
    'Serie': {
        'value': None,
        'required': True,
        'position_original_document': 4
    },
    'Numar': {
        'value': None,
        'required': True,
        'position_original_document': 5
    },
    'Punct de lucru': {
        'value': 'sediu',
        'required': True,
        'position_original_document': None
    },
    'Data': {
        'value': None,
        'required': False,
        'position_original_document': 2
    },
    'Data scadenta': {
        'value': None,
        'required': False,
        'position_original_document': 2
    },
    'Cod partener': {
        'value': None,
        'required': False,
        'position_original_document': 9
    },
    'Nume partener': {
        'value': None,
        'required': False,
        'position_original_document': 10
    },
    'Atribut fiscal': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Cod fiscal': {
        'value': None,
        'required': False,
        'position_original_document': 10
    },
    'Nr.Reg.Com.': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Pct. Lucru partener': {
        'value': 'Sediu',
        'required': False,
        'position_original_document': None
    },
    'Rezidenta': {
        'value': 'Romania',
        'required': False,
        'position_original_document': None
    },
    'Tara': {
        'value': 'Romania',
        'required': False,
        'position_original_document': None
    },
    'Judet': {
        'value': 'Dolj',
        'required': False,
        'position_original_document': None
    },
    'Localitate': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Strada': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Numar': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Bloc': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Scara': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Etaj': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Apartament': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Cod postal': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Moneda': {
        'value': "RON",
        'required': False,
        'position_original_document': None
    },
    'Curs': {
        'value': "1.00",
        'required': False,
        'position_original_document': None
    },
    'TVA la incasare': {
        'value': "1",
        'required': False,
        'position_original_document': None
    },
    'Taxare inversa': {
        'value': "0",
        'required': False,
        'position_original_document': None
    },
    'Cod agent': {
        'value': None,
        'required': False,
        'position_original_document': None
    },
    'Valoare neta totala': {
        'value': None,
        'required': True,
        'position_original_document': None
    },
    'Valoare TVA': {
        'value': None,
        'required': True,
        'position_original_document': None
    },
    'Total document': {
        'value': None,
        'required': True,
        'position_original_document': None
    }
}

VAT_POSITIONS = {
    '19': {
        'total_amount': 17,
        'vat': 18
    },
    '9': {
        'total_amount': 19,
        'vat': 20
    }
}

ROW_START = 12
ROW_ITEM_COUNT = 30
ROW_GAP = 17