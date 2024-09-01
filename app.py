import os
import openai
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Setup OpenAI
os.environ['OPENAI_API_KEY'] = 'fake_key'
os.environ['SPREAD_SHEETS_ID'] = '1-9IRLgu7nqAnM7fJ6kaiKSCA3sodmqt5UgZ9iCBToKE'
openai.api_key = os.getenv('OPENAI_API_KEY')
client = openai.OpenAI(
    api_key=openai.api_key,
    base_url="http://localhost:11434/v1"
)

# Setup Google Sheets
spread_sheets_id = os.getenv('SPREAD_SHEETS_ID')
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('/Users/anas1/Downloads/key.json', scope)
gc = gspread.authorize(credentials)
sheet = gc.open_by_key(spread_sheets_id).sheet1


# Classify comments function
def classify_comment(comment):
    try:
        response = client.chat.completions.create(
            model="llama2",
            
            messages=[
                {"role": "system", "content": """
                    You are a helpful assistant designed to classify comments into: 
                    Neutral, Request, Question, Appreciation, Error or Other. 
                    Must respond with one word.
                    """
                },
                {"role": "user", "content": "Great product, very satisfied with the purchase."},
                {"role": "assistant", "content": "Appreciation"},
                {"role": "user", "content": "Can you provide more details about this product?"},
                {"role": "assistant", "content": "Question"},
                {"role": "user", "content": "Please create a product on how to use the new software."},
                {"role": "assistant", "content": "Request"},
                {"role": "user", "content": "The service was terrible, I won't be coming back."},
                {"role": "assistant", "content": "Error"},
                {"role": "user", "content": "This product is okay, nothing special."},
                {"role": "assistant", "content": "Neutral"},
                {"role": "user", "content": comment}  # This is the comment you want to classify
            ]
        )
        
        category = response.choices[0].message.content.strip()
        print(f"Comment: '{comment}' -> Category: '{category}'")
        return category.title()
    except Exception as e:
        print(f"Error in classifying comment: {e}")
        return "Process Failed"
    
# Read data from the sheet
data = sheet.get_all_values()
header = data[0]  # Assuming the first row is the header
comments_rows = data[1:]  # Exclude header
# Iterate over each row that contains comments
for i, row in enumerate(comments_rows, start=2):  # Google Sheets is 1-indexed; header is row 1, so data starts at row 2
    comment = row[0]  # Assuming the comment is in the second column
    category = classify_comment(comment)
    # Update the 3th column (index 2) with the classification
    sheet.update_cell(i, 2, category)

print("The Google Sheet has been updated with categories for each comment.")

