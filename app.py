import os
from dotenv import load_dotenv
from flask import Flask, render_template, request, jsonify
from excel_ai_agent import ExcelAIAgent

# Load environment variables
load_dotenv()

# Initialize Flask app
app = Flask(__name__)

# Check if the API key is available
gemini_api_key = os.getenv("GEMINI_API_KEY")
if not gemini_api_key:
    print("Warning: GEMINI_API_KEY environment variable not found.")
    print("Please set it in a .env file or in your environment.")

# Initialize Excel AI Agent
excel_agent = ExcelAIAgent()

@app.route('/')
def index():
    """Render the main page of the application."""
    return render_template('index.html')

@app.route('/api/query', methods=['POST'])
def process_query():
    """Process a query from the user and return the AI response."""
    if not request.json or 'query' not in request.json:
        return jsonify({'error': 'Invalid request'}), 400
    
    user_query = request.json['query']
    excel_context = request.json.get('context', {})
    image_data = request.json.get('image_data')
    
    try:
        response = excel_agent.process_query(user_query, excel_context, image_data)
        return jsonify({'response': response})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/autonomous', methods=['POST'])
def autonomous_action():
    """Process a query and autonomously execute actions in Excel."""
    if not request.json or 'query' not in request.json:
        return jsonify({'error': 'Invalid request'}), 400
    
    user_query = request.json['query']
    excel_context = request.json.get('context', {})
    image_data = request.json.get('image_data')
    
    try:
        result = excel_agent.process_autonomous_query(user_query, excel_context, image_data)
        return jsonify(result)
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e),
            'action_taken': None,
            'explanation': f"An unexpected error occurred: {str(e)}"
        }), 500

@app.route('/api/connect', methods=['POST'])
def connect_to_excel():
    """Connect to an Excel instance."""
    try:
        connection_status = excel_agent.connect()
        return jsonify({'status': connection_status})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/execute', methods=['POST'])
def execute_command():
    """Execute a command in Excel."""
    if not request.json or 'command' not in request.json:
        return jsonify({'error': 'Invalid request'}), 400
    
    command = request.json['command']
    
    try:
        result = excel_agent.execute_command(command)
        return jsonify({'result': result})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/health', methods=['GET'])
def health_check():
    """Check if the API is running and Excel is connected."""
    # Safely check Excel connection status
    try:
        excel_connected = excel_agent.excel_app is not None and excel_agent.workbook is not None
    except:
        excel_connected = False
    
    gemini_configured = excel_agent.api_key is not None
    
    response = {
        'status': 'ok',
        'excel_connected': excel_connected,
        'gemini_configured': gemini_configured
    }
    
    # Add workbook name if connected - safely
    if excel_connected:
        try:
            workbook_name = excel_agent.workbook.name
            response['workbook_name'] = workbook_name
        except Exception as e:
            # If we can't get the name, don't fail the whole request
            response['workbook_name'] = "Unknown"
            response['excel_connection_issue'] = str(e)
    
    return jsonify(response)

if __name__ == '__main__':
    app.run(debug=True) 