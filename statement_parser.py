import pandas as pd
import re
from datetime import datetime
import tabula  # For PDF parsing
import pytesseract  # For image/screenshot parsing
from PIL import Image
import cv2
import numpy as np

class StatementParser:
    def __init__(self):
        # Common patterns in Santander statements
        self.date_patterns = [
            r'\d{2}/\d{2}/\d{4}',  # DD/MM/YYYY
            r'\d{2}-\d{2}-\d{4}',  # DD-MM-YYYY
            r'\d{2}\s+[A-Za-z]{3}\s+\d{4}'  # DD MMM YYYY
        ]
        
        # Common expense patterns
        self.expense_patterns = [
            r'CARD PAYMENT TO.*?(\d+\.\d{2})',
            r'DIRECT DEBIT.*?(\d+\.\d{2})',
            r'FASTER PAYMENT TO.*?(\d+\.\d{2})',
            r'ATM WITHDRAWAL.*?(\d+\.\d{2})',
            r'STANDING ORDER TO.*?(\d+\.\d{2})'
        ]
        
        # Common income patterns
        self.income_patterns = [
            r'FASTER PAYMENT FROM.*?(\d+\.\d{2})',
            r'DEPOSIT.*?(\d+\.\d{2})',
            r'SALARY.*?(\d+\.\d{2})'
        ]

    def parse_pdf(self, pdf_path):
        """Parse PDF bank statement"""
        try:
            # Read PDF file
            tables = tabula.read_pdf(pdf_path, pages='all')
            
            # Combine all tables
            df = pd.concat(tables, ignore_index=True)
            
            # Process the dataframe
            return self._process_statement_data(df)
            
        except Exception as e:
            print(f"Error parsing PDF: {str(e)}")
            return None

    def parse_image(self, image_path):
        """Parse image/screenshot of bank statement"""
        try:
            # Read image
            image = cv2.imread(image_path)
            
            # Preprocess image
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
            
            # Extract text using OCR
            text = pytesseract.image_to_string(thresh)
            
            # Process the text
            return self._process_text_statement(text)
            
        except Exception as e:
            print(f"Error parsing image: {str(e)}")
            return None

    def _process_text_statement(self, text):
        """Process text extracted from statement"""
        transactions = []
        
        # Split into lines
        lines = text.split('\n')
        
        for line in lines:
            # Skip empty lines
            if not line.strip():
                continue
                
            # Try to find date
            date = None
            for pattern in self.date_patterns:
                match = re.search(pattern, line)
                if match:
                    date_str = match.group()
                    try:
                        if '/' in date_str:
                            date = datetime.strptime(date_str, '%d/%m/%Y')
                        elif '-' in date_str:
                            date = datetime.strptime(date_str, '%d-%m-%Y')
                        else:
                            date = datetime.strptime(date_str, '%d %b %Y')
                        break
                    except ValueError:
                        continue
            
            if not date:
                continue
                
            # Try to find amount and determine if expense or income
            amount = None
            is_expense = False
            
            # Check expense patterns
            for pattern in self.expense_patterns:
                match = re.search(pattern, line)
                if match:
                    amount = float(match.group(1))
                    is_expense = True
                    break
            
            # Check income patterns if no expense found
            if amount is None:
                for pattern in self.income_patterns:
                    match = re.search(pattern, line)
                    if match:
                        amount = float(match.group(1))
                        break
            
            if amount is not None:
                # Get description (remove date and amount)
                description = line
                for pattern in self.date_patterns + [r'\d+\.\d{2}']:
                    description = re.sub(pattern, '', description)
                description = description.strip()
                
                transactions.append({
                    'date': date.strftime('%Y-%m-%d'),
                    'description': description,
                    'cost': -amount if is_expense else amount
                })
        
        return pd.DataFrame(transactions)

    def _process_statement_data(self, df):
        """Process dataframe from PDF parsing"""
        # Try to identify date and amount columns
        date_col = None
        amount_col = None
        desc_col = None
        
        for col in df.columns:
            # Look for date column
            if df[col].astype(str).str.match(self.date_patterns[0]).any():
                date_col = col
            # Look for amount column (contains currency symbols or decimal numbers)
            elif df[col].astype(str).str.contains(r'£?\d+\.\d{2}').any():
                amount_col = col
            # Assume the longest text column is description
            elif df[col].astype(str).str.len().mean() > 20:
                desc_col = col
        
        if not all([date_col, amount_col, desc_col]):
            return None
            
        # Create new dataframe with standardized columns
        transactions = []
        
        for _, row in df.iterrows():
            try:
                # Parse date
                date_str = str(row[date_col])
                for pattern in self.date_patterns:
                    match = re.search(pattern, date_str)
                    if match:
                        date_str = match.group()
                        break
                
                if '/' in date_str:
                    date = datetime.strptime(date_str, '%d/%m/%Y')
                elif '-' in date_str:
                    date = datetime.strptime(date_str, '%d-%m-%Y')
                else:
                    date = datetime.strptime(date_str, '%d %b %Y')
                
                # Parse amount
                amount_str = str(row[amount_col])
                amount_match = re.search(r'£?(\d+\.\d{2})', amount_str)
                if amount_match:
                    amount = float(amount_match.group(1))
                    # Determine if expense based on description
                    is_expense = any(re.search(pattern, str(row[desc_col])) for pattern in self.expense_patterns)
                    
                    transactions.append({
                        'date': date.strftime('%Y-%m-%d'),
                        'description': str(row[desc_col]).strip(),
                        'cost': -amount if is_expense else amount
                    })
            except (ValueError, AttributeError):
                continue
        
        return pd.DataFrame(transactions)

    def suggest_categories(self, transactions_df):
        """Suggest categories based on transaction descriptions"""
        category_patterns = {
            'Groceries': [
                r'TESCO', r'SAINSBURY', r'ASDA', r'ALDI', r'LIDL', r'MORRISONS',
                r'WAITROSE', r'CO-OP', r'FOOD', r'GROCERY'
            ],
            'Transportation': [
                r'TRANSPORT', r'TFL', r'TRAIN', r'BUS', r'UBER', r'TAXI',
                r'PARKING', r'FUEL', r'PETROL', r'SHELL', r'BP', r'ESSO'
            ],
            'Entertainment': [
                r'CINEMA', r'NETFLIX', r'SPOTIFY', r'AMAZON PRIME',
                r'THEATRE', r'TICKET', r'GAME', r'STEAM'
            ],
            'Bills & Utilities': [
                r'WATER', r'ELECTRIC', r'GAS', r'ENERGY', r'COUNCIL TAX',
                r'PHONE', r'MOBILE', r'INTERNET', r'BROADBAND', r'TV LICENSE'
            ],
            'Dining Out': [
                r'RESTAURANT', r'CAFE', r'COFFEE', r'STARBUCKS',
                r'COSTA', r'MCDONALDS', r'KFC', r'TAKEAWAY', r'DELIVEROO',
                r'JUST EAT', r'UBER EATS'
            ],
            'Shopping': [
                r'AMAZON', r'EBAY', r'ARGOS', r'BOOTS', r'SUPERDRUG',
                r'NEXT', r'PRIMARK', r'H&M', r'ASOS'
            ],
            'Income': [
                r'SALARY', r'DEPOSIT', r'FASTER PAYMENT FROM'
            ]
        }
        
        def suggest_category(description):
            description = description.upper()
            for category, patterns in category_patterns.items():
                if any(re.search(pattern, description, re.IGNORECASE) for pattern in patterns):
                    return category
            return 'Other'
        
        transactions_df['suggested_category'] = transactions_df['description'].apply(suggest_category)
        return transactions_df 