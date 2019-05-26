index = "transaction_number"


def rows_to_skip(x):
    values_to_skip = ['ID', 'transaction_id', 'TransactionID', 'ID', 'TRNX ID', 'ID', 'TRANSACTION ID', 'Card Serno',
                      'Transaction ID', 'Source Reg Num', 'id', 'ID of the transaction', 'transaction_number']
    if x in values_to_skip:
        return True
    return False

