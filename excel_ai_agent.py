import os
import json
import google.generativeai as genai
import xlwings as xw
import pandas as pd
import numpy as np
import re
from typing import Dict, Any, List, Optional, Tuple, Union
import time
import win32com.client
import pythoncom
import logging
from google.generativeai.types import Part

class ExcelAIAgent:
    """
    AI-powered agent for Microsoft Excel that can perform various tasks using the Gemini API.
    """
    
    def __init__(self):
        """Initialize the Excel AI Agent"""
        # Version information
        self.version = "1.0.0"
        
        # Excel objects
        self.excel_app = None
        self.workbook = None
        self.sheet = None
        
        # API configuration
        self.api_key = os.environ.get("GEMINI_API_KEY")
        if not self.api_key:
            raise ValueError("GEMINI_API_KEY environment variable not set")
        
        # Configure the Gemini API
        genai.configure(api_key=self.api_key)
        self.model = genai.GenerativeModel('gemini-1.5-pro')
        
        # Excel constants
        self.excel_constants = {
            "chart_types": {
                "column": 3,  # xlColumnClustered
                "bar": 57,     # xlBarClustered
                "line": 4,     # xlLine
                "pie": 5,      # xlPie
                "scatter": 74, # xlXYScatterSmooth
                "area": 76     # xlAreaStacked
            }
        }
        
        # Define system prompts for different tasks
        self.system_prompts = {
            "general": """You are an autonomous Excel AI Assistant that directly performs tasks in Microsoft Excel. 
                       You not only provide answers to Excel-related questions but also execute tasks on behalf of the user.
                       Analyze the user's request and determine what actions to take in Excel.
                       For each task, extract the specific commands that need to be executed in Excel and explain what you are doing.
                       When creating formulas or automating tasks, explain how they work in plain language.
                       If the user provides an image, analyze it to understand the Excel content, layout, data, or error messages shown in it.""",
                       
            "formula": """You are an autonomous Excel formula expert that automatically implements Excel formulas.
                       Based on the user's description, create and apply the appropriate formula to the specified cells.
                       Explain the formula's purpose, structure, and logic before implementing it.
                       Always extract the target cell or range for the formula and provide a clear implementation plan.
                       If an image of a spreadsheet or formula is provided, analyze it to understand the current structure,
                       data relationships, or any formula errors that need to be addressed.""",
                       
            "analysis": """You are an autonomous Excel data analyst that automatically analyzes spreadsheet data.
                       Examine the data provided, identify patterns, anomalies, and insights.
                       Automatically generate appropriate visualizations and summary statistics.
                       Format data appropriately to highlight important findings.
                       Provide your analysis in an easy-to-understand format while explaining what actions you're taking.
                       If an image of data or charts is provided, analyze it to understand the data structure, 
                       trends, and insights that can be observed.""",
                       
            "troubleshooting": """You are an autonomous Excel troubleshooter that automatically identifies and fixes issues.
                               Diagnose problems in formulas, data structures, or formatting.
                               Automatically correct errors when possible and explain the solutions.
                               For complex issues, provide step-by-step repair instructions and implement them when authorized.
                               If an image showing Excel errors or issues is provided, analyze it carefully to identify
                               the specific error type, its cause, and the most appropriate solution.""",
                               
            "automation": """You are an autonomous Excel automation specialist that implements time-saving processes.
                          Analyze the user's workflow request and implement the appropriate automation solution.
                          Create and apply macros, formulas, or data transformations to automate repetitive tasks.
                          Optimize spreadsheet structure for efficiency while explaining your implementation approach.
                          If an image of a workflow or spreadsheet structure is provided, analyze it to understand
                          the current process and identify opportunities for automation."""
        }
        
        # Command hints for the AI to recognize and extract actionable tasks
        self.command_patterns = {
            "insert_data": [r"insert\s+(.+?)\s+into\s+(.+)", r"put\s+(.+?)\s+in\s+(.+)", r"add\s+(.+?)\s+to\s+(.+)"],
            "create_formula": [r"create\s+formula\s+(.+?)\s+in\s+(.+)", r"calculate\s+(.+?)\s+in\s+(.+)", r"compute\s+(.+?)\s+in\s+(.+)"],
            "format_cells": [r"format\s+(.+?)\s+as\s+(.+)", r"set\s+(.+?)\s+format\s+to\s+(.+)", r"style\s+(.+?)\s+with\s+(.+)"],
            "create_chart": [r"create\s+(.+?)\s+chart\s+from\s+(.+)", r"graph\s+(.+?)\s+using\s+(.+)", r"visualize\s+(.+?)\s+with\s+(.+)"],
            "sort_data": [r"sort\s+(.+?)\s+by\s+(.+)", r"order\s+(.+?)\s+using\s+(.+)", r"arrange\s+(.+?)\s+by\s+(.+)"],
            "filter_data": [r"filter\s+(.+?)\s+by\s+(.+)", r"show\s+only\s+(.+?)\s+from\s+(.+)", r"find\s+(.+?)\s+in\s+(.+)"]
        }
    
    def connect(self):
        """Connect to Excel"""
        try:
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Use xlwings to connect to Excel instead of direct win32com calls
            try:
                # Try to connect to active Excel instance
                self.excel_app = xw.apps.active
                print("Connected to existing Excel instance via xlwings")
            except:
                # If no active instance, create a new one
                self.excel_app = xw.App(visible=True)
                print("Created new Excel instance via xlwings")
            
            # Make Excel visible and interactive
            self.excel_app.visible = True
            self.excel_app.screen_updating = True
            
            # Get active workbook or create one
            try:
                # Try to get the active workbook
                workbooks = self.excel_app.books
                if len(workbooks) > 0:
                    self.workbook = workbooks.active
                    print(f"Connected to active workbook: {self.workbook.name}")
                else:
                    # Create a new workbook if none exists
                    self.workbook = self.excel_app.books.add()
                    print("Created new workbook via xlwings")
            except Exception as e:
                print(f"Error getting or creating workbook: {str(e)}")
                # Try creating a new workbook as fallback
                try:
                    self.workbook = self.excel_app.books.add()
                    print("Created new workbook as fallback")
                except Exception as wb_err:
                    print(f"Failed to create workbook: {str(wb_err)}")
                    self.workbook = None
            
            # Get active sheet or create one
            if self.workbook:
                try:
                    # Get the active sheet
                    self.sheet = self.workbook.sheets.active
                    print(f"Connected to active sheet: {self.sheet.name}")
                except Exception as active_sheet_err:
                    print(f"Error getting active sheet: {str(active_sheet_err)}")
                    # Try getting the first sheet
                    try:
                        if len(self.workbook.sheets) > 0:
                            self.sheet = self.workbook.sheets[0]
                            print(f"Using first sheet: {self.sheet.name}")
                        else:
                            # Create a new sheet if none exists
                            self.sheet = self.workbook.sheets.add()
                            print("Created new sheet via xlwings")
                    except Exception as sheet_err:
                        print(f"Error getting or creating sheet: {str(sheet_err)}")
                        self.sheet = None
            
            if self.workbook and self.sheet:
                return {"status": "success", "message": "Connected to Excel with active sheet"}
            elif self.workbook:
                return {"status": "partial", "message": "Connected to Excel but no active sheet"}
            else:
                return {"status": "error", "message": "Connected to Excel but couldn't access a workbook"}
        
        except Exception as e:
            print(f"Error connecting to Excel: {str(e)}")
            self.excel_app = None
            self.workbook = None
            self.sheet = None
            return {"status": "error", "message": f"Failed to connect to Excel: {str(e)}"}
    
    def process_query(self, query: str, context: Dict[str, Any] = None, image_data: str = None) -> str:
        """
        Process a user query about Excel and return an AI-generated response.
        
        Args:
            query: The user's question or request
            context: Additional context about the current Excel state
            image_data: Optional base64-encoded image data
            
        Returns:
            str: AI-generated response to the query
        """
        if not self.api_key:
            return "Gemini API key not configured. Please set the GEMINI_API_KEY environment variable."
        
        # Determine the type of query to select appropriate system prompt
        query_type = self._determine_query_type(query)
        system_prompt = self.system_prompts.get(query_type, self.system_prompts["general"])
        
        # Get current Excel context if available and connected
        excel_context = self._get_excel_context() if self.excel_app else {}
        
        # Combine user provided context with excel context
        if context:
            excel_context.update(context)
        
        # Build prompt with context
        prompt = self._build_prompt(query, system_prompt, excel_context)
        
        # Call Gemini API with or without image
        if image_data:
            # Process the image data
            image_parts = self._process_image_data(image_data)
            response = self.model.generate_content([prompt, image_parts])
        else:
            response = self.model.generate_content(prompt)
        
        return response.text
    
    def _determine_query_type(self, query: str) -> str:
        """
        Determine the type of Excel query to select the appropriate system prompt.
        
        Args:
            query: The user's question or request
            
        Returns:
            str: The determined query type
        """
        query = query.lower()
        
        if any(word in query for word in ["automate", "automation", "macro", "automatically", "batch", "repeat"]):
            return "automation"
        elif any(word in query for word in ["formula", "function", "calculation", "compute", "='", "=", "calculate"]):
            return "formula"
        elif any(word in query for word in ["analyze", "analysis", "pattern", "trend", "insight", "dashboard", "chart", "graph"]):
            return "analysis"
        elif any(word in query for word in ["error", "issue", "problem", "fix", "troubleshoot", "not working", "broken"]):
            return "troubleshooting"
        else:
            return "general"
    
    def _build_prompt(self, query: str, system_prompt: str, context: Dict[str, Any]) -> str:
        """
        Build a prompt for the Gemini API with system prompt and context.
        
        Args:
            query: The user's question or request
            system_prompt: The system prompt for the specific query type
            context: Additional context about the current Excel state
            
        Returns:
            str: The complete prompt for the Gemini API
        """
        # Start with the system prompt
        full_prompt = f"{system_prompt}\n\n"
        
        # Add context if available
        if context:
            context_str = json.dumps(context, indent=2)
            full_prompt += f"Excel Context:\n{context_str}\n\n"
        
        # Add the user query
        full_prompt += f"User Query: {query}"
        
        return full_prompt
    
    def _get_excel_context(self) -> Dict[str, Any]:
        """
        Get context from the current Excel workbook.
        
        Returns:
            dict: Context information from Excel
        """
        # Initialize COM for this thread
        try:
            pythoncom.CoInitialize()
        except:
            pass
            
        if not self.excel_app:
            return {"error": "Not connected to Excel"}
            
        # Reconnect if needed
        if not self.workbook or not self.sheet:
            try:
                self.connect()
                if not self.workbook or not self.sheet:
                    return {"error": "Could not reconnect to Excel workbook or sheet"}
            except Exception as e:
                return {"error": f"Error reconnecting to Excel: {str(e)}"}
        
        try:
            # Get basic workbook and sheet information
            workbook_name = self.workbook.name
            sheet_name = self.sheet.name
            
            # Get context data
            context = {
                "workbook_name": workbook_name,
                "sheet_name": sheet_name,
                "selection": self._get_selection_info_safe(),
                "sheet_data": self._get_sheet_data_sample_safe(),
                "formulas": self._get_formulas_in_use_safe()
            }
            return context
        except Exception as e:
            print(f"Error getting Excel context: {str(e)}")
            return {"error": f"Error getting Excel context: {str(e)}"}
    
    def _get_selection_info_safe(self) -> Dict[str, Any]:
        """
        Safely get information about the current selection in Excel.
        
        Returns:
            dict: Information about the current selection
        """
        try:
            # Get the current selection using xlwings
            selection = self.excel_app.selection
            address = selection.address
            value = selection.value
            
            return {
                "address": address,
                "value": str(value) if value is not None else None
            }
        except Exception as e:
            print(f"Error getting selection info: {str(e)}")
            return {"error": "Could not retrieve selection information"}
    
    def _get_sheet_data_sample_safe(self, max_rows: int = 10, max_cols: int = 10) -> List[List[Any]]:
        """
        Safely get a sample of data from the current sheet.
        
        Args:
            max_rows: Maximum number of rows to include
            max_cols: Maximum number of columns to include
            
        Returns:
            list: A 2D list containing the sampled data
        """
        try:
            # Get the used range using xlwings
            used_range = self.sheet.used_range
            
            # Determine the range to sample
            last_row = min(used_range.last_cell.row, max_rows)
            last_col = min(used_range.last_cell.column, max_cols)
            
            # Get the data
            sample_range = self.sheet.range((1, 1), (last_row, last_col))
            data = sample_range.value
            
            # Ensure we have a 2D list
            if not isinstance(data, list):
                data = [[data]]
            elif data and not isinstance(data[0], list):
                data = [data]
            
            return data
        except Exception as e:
            print(f"Error getting sheet data: {str(e)}")
            # Return empty data on error
            return [["No data available"]]
    
    def _get_formulas_in_use_safe(self, max_formulas: int = 10) -> List[Dict[str, str]]:
        """
        Safely get formulas currently in use in the sheet.
        
        Args:
            max_formulas: Maximum number of formulas to include
            
        Returns:
            list: A list of dictionaries containing formula addresses and formula strings
        """
        try:
            # Get the used range
            used_range = self.sheet.used_range
            
            # Create a list to store formulas
            formulas = []
            formula_count = 0
            
            # Iterate through rows and columns looking for formulas
            for row in range(1, used_range.last_cell.row + 1):
                if formula_count >= max_formulas:
                    break
                    
                for col in range(1, used_range.last_cell.column + 1):
                    cell = self.sheet.range((row, col))
                    formula = cell.formula
                    
                    if formula and formula.startswith('='):
                        formulas.append({
                            "address": cell.address,
                            "formula": formula
                        })
                        
                        formula_count += 1
                        if formula_count >= max_formulas:
                            break
            
            return formulas
        except Exception as e:
            print(f"Error getting formulas: {str(e)}")
            return []
    
    def _process_image_data(self, image_data: str):
        """
        Process the base64 image data for use with Gemini API.
        
        Args:
            image_data: Base64-encoded image data
            
        Returns:
            Part: The image part for Gemini API
        """
        # Strip the data URL prefix if present
        if image_data.startswith('data:image'):
            # Extract the base64 part
            image_data = image_data.split(',', 1)[1]
        
        # Decode base64 data
        import base64
        from google.generativeai.types import Part
        
        image_bytes = base64.b64decode(image_data)
        return Part.from_bytes(image_bytes, mime_type="image/jpeg")
    
    def process_autonomous_query(self, query: str, context: Dict[str, Any] = None, image_data: str = None) -> Dict[str, Any]:
        """
        Process and execute tasks based on a user query.
        
        Args:
            query: The user's task request
            context: Additional context about the current Excel state
            image_data: Optional base64-encoded image data
            
        Returns:
            dict: Result of the operation with action taken and explanation
        """
        # Check Excel connection and attempt to reconnect if needed
        if not self.excel_app or not self.workbook or not self.sheet:
            try:
                connection_status = self.connect()
                # If still no sheet, try to create one
                if not self.sheet and self.workbook:
                    try:
                        try:
                            # Try to access the first sheet directly
                            self.sheet = self.workbook.Worksheets(1)
                            print("Found first worksheet in autonomous query")
                        except:
                            # Create a new sheet if can't access any existing one
                            self.sheet = self.workbook.Worksheets.Add()
                            print("Created new worksheet in autonomous query")
                    except Exception as sheet_err:
                        print(f"Error creating new worksheet: {sheet_err}")
                
                if not self.sheet:
                    return {
                        "status": "error",
                        "message": "Could not access or create a worksheet in Excel.",
                        "action_taken": None,
                        "explanation": "No action was taken because a worksheet could not be accessed or created in Excel."
                    }
                    
                if "error" in connection_status.get("status", ""):
                    return {
                        "status": "error",
                        "message": "Not connected to Excel. Please connect first.",
                        "action_taken": None,
                        "explanation": f"No action was taken because there is no active Excel connection. {connection_status.get('message', '')}"
                    }
            except Exception as e:
                return {
                    "status": "error",
                    "message": f"Excel connection error: {str(e)}",
                    "action_taken": None,
                    "explanation": "No action was taken because there was an error connecting to Excel."
                }
        
        if not self.api_key:
            return {
                "status": "error", 
                "message": "Gemini API key not configured.",
                "action_taken": None,
                "explanation": "No action was taken because the AI model could not be accessed."
            }
        
        # Verify Excel connection is still valid
        try:
            # Simple check to see if Excel is still responsive
            # Excel app doesn't have a 'name' attribute, so check workbook instead
            test = self.workbook.name
            test_sheet = self.sheet.name
        except Exception as e:
            # Try to reconnect once
            try:
                connection_status = self.connect()
                if "Failed" in connection_status:
                    return {
                        "status": "error",
                        "message": "Excel connection lost. Please reconnect.",
                        "action_taken": None,
                        "explanation": f"No action was taken because the Excel connection was lost. {connection_status}"
                    }
            except Exception as ex:
                return {
                    "status": "error",
                    "message": f"Excel connection error: {str(ex)}",
                    "action_taken": None,
                    "explanation": "No action was taken because there was an error with the Excel connection."
                }
        
        try:
            # Get current Excel context
            excel_context = self._get_excel_context()
            
            # Check if there was an error getting context
            if "error" in excel_context:
                return {
                    "status": "error",
                    "message": "Error accessing Excel context",
                    "action_taken": None,
                    "explanation": f"No action was taken because: {excel_context['error']}"
                }
            
            if context:
                excel_context.update(context)
            
            # Determine query type
            query_type = self._determine_query_type(query)
            
            # Get action plan from AI
            action_plan = self._get_action_plan(query, query_type, excel_context, image_data)
            
            # Extract and execute commands
            commands = self._extract_commands_from_plan(action_plan)
            
            if not commands:
                return {
                    "status": "warning",
                    "message": "No actionable commands were extracted from the query.",
                    "action_taken": [],
                    "explanation": "I understood your request but couldn't determine specific Excel actions to take. Please try rephrasing your request with more specific instructions."
                }
            
            # Execute the commands
            results = []
            for cmd in commands:
                result = self.execute_command(cmd)
                results.append(result)
            
            # Create a summary of what was done
            summary = self._create_execution_summary(query, action_plan, commands, results)
            
            return {
                "status": "success" if all(r.get("status") == "success" for r in results) else "partial_success",
                "message": summary["message"],
                "action_taken": summary["actions"],
                "explanation": summary["explanation"]
            }
        except Exception as e:
            return {
                "status": "error",
                "message": f"Error processing query: {str(e)}",
                "action_taken": None,
                "explanation": f"An unexpected error occurred while processing your request: {str(e)}"
            }
    
    def _get_action_plan(self, query: str, query_type: str, context: Dict[str, Any], image_data: str = None) -> str:
        """
        Get an action plan from the AI model based on the user query.
        
        Args:
            query: The user's request
            query_type: The type of query
            context: Excel context information
            image_data: Optional base64-encoded image data
            
        Returns:
            str: The AI-generated action plan
        """
        system_prompt = self.system_prompts.get(query_type, self.system_prompts["general"])
        
        # Build a more specific prompt for action planning
        action_prompt = f"{system_prompt}\n\n"
        action_prompt += "You need to extract actionable Excel commands from user requests.\n"
        action_prompt += "For each request, provide:\n"
        action_prompt += "1. A brief explanation of what will be done\n"
        action_prompt += "2. A list of specific Excel commands to execute\n"
        action_prompt += "3. The expected outcome\n\n"
        
        if context:
            context_str = json.dumps(context, indent=2)
            action_prompt += f"Excel Context:\n{context_str}\n\n"
        
        action_prompt += f"User Request: {query}\n\n"
        action_prompt += "Please create an action plan with commands formatted as JSON that can be executed directly in Excel."
        
        # Call Gemini API with or without image
        if image_data:
            # Process the image data
            image_parts = self._process_image_data(image_data)
            response = self.model.generate_content([action_prompt, image_parts])
        else:
            response = self.model.generate_content(action_prompt)
        
        return response.text
    
    def _extract_commands_from_plan(self, action_plan: str) -> List[Dict[str, Any]]:
        """
        Extract executable commands from the AI-generated action plan.
        
        Args:
            action_plan: The AI-generated action plan
            
        Returns:
            list: List of command dictionaries to execute
        """
        commands = []
        
        # Try to extract JSON formatted commands
        try:
            # Look for JSON code blocks
            json_pattern = r"```(?:json)?\s*(\{.*?\})\s*```"
            matches = re.findall(json_pattern, action_plan, re.DOTALL)
            
            if matches:
                for match in matches:
                    try:
                        cmd = json.loads(match)
                        if isinstance(cmd, dict) and "type" in cmd:
                            commands.append(cmd)
                        elif isinstance(cmd, list):
                            for c in cmd:
                                if isinstance(c, dict) and "type" in c:
                                    commands.append(c)
                    except json.JSONDecodeError:
                        continue
        except Exception:
            pass
        
        # If no JSON commands found, try to extract using command patterns
        if not commands:
            for cmd_type, patterns in self.command_patterns.items():
                for pattern in patterns:
                    matches = re.findall(pattern, action_plan, re.IGNORECASE)
                    for match in matches:
                        if len(match) >= 2:
                            cmd = {
                                "type": cmd_type,
                                "parameters": list(match)
                            }
                            commands.append(cmd)
        
        # If still no commands, create a generic command based on query type detection
        if not commands:
            # Extract cell references from the plan
            cell_ref_pattern = r'([A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?)'
            cell_refs = re.findall(cell_ref_pattern, action_plan)
            
            # Extract formulas from the plan
            formula_pattern = r'=([^=\n]+)'
            formulas = re.findall(formula_pattern, action_plan)
            
            if cell_refs and formulas:
                commands.append({
                    "type": "insert_formula",
                    "cell": cell_refs[0],
                    "formula": f"={formulas[0].strip()}"
                })
        
        return commands
    
    def _create_execution_summary(self, query: str, action_plan: str, commands: List[Dict[str, Any]], results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Create a summary of actions taken based on the original query and results.
        
        Args:
            query: The original user query
            action_plan: The AI-generated action plan
            commands: The commands that were executed
            results: The results of command execution
            
        Returns:
            dict: Summary information
        """
        # Count successful commands
        success_count = sum(1 for r in results if r.get("status") == "success")
        
        # Extract explanations from the action plan
        explanation_pattern = r"(?:explanation|I will|Here's what I'll do):(.*?)(?=\n\n|\Z)"
        explanation_matches = re.findall(explanation_pattern, action_plan, re.IGNORECASE | re.DOTALL)
        explanation = explanation_matches[0].strip() if explanation_matches else "Executed the requested Excel operations."
        
        # Create actions list
        actions = []
        for cmd, result in zip(commands, results):
            cmd_type = cmd.get("type", "unknown")
            cmd_status = result.get("status", "unknown")
            cmd_message = result.get("message", "")
            
            if cmd_type == "insert_formula":
                action = f"Inserted formula {cmd.get('formula', '')} at {cmd.get('cell', '')}"
            elif cmd_type == "format_range":
                action = f"Formatted range {cmd.get('range', '')}"
            elif cmd_type == "create_chart":
                action = f"Created {cmd.get('chart_type', 'chart')} using data from {cmd.get('data_range', '')}"
            elif cmd_type == "get_data":
                action = f"Retrieved data from range {cmd.get('range', '')}"
            else:
                action = f"Executed {cmd_type} command"
                
            if cmd_status != "success":
                action += f" (Failed: {result.get('error', 'Unknown error')})"
                
            actions.append(action)
            
        if not actions:
            actions = ["Analyzed the request but no specific Excel operations were needed."]
            
        # Create a message
        if success_count == len(commands) and commands:
            message = f"Successfully executed {success_count} Excel operations."
        elif success_count > 0:
            message = f"Partially completed the request. {success_count} out of {len(commands)} operations succeeded."
        else:
            message = "Could not execute the requested operations in Excel."
            
        return {
            "message": message,
            "actions": actions,
            "explanation": explanation
        }
    
    def execute_command(self, command: Dict[str, Any]) -> Dict[str, Any]:
        """
        Execute a command in Excel.
        
        Args:
            command: A dictionary describing the command to execute
            
        Returns:
            dict: The result of the command execution
        """
        # Ensure COM is initialized for this thread
        try:
            pythoncom.CoInitialize()
        except:
            pass
            
        # Verify Excel connection is still valid
        if not self.excel_app or not self.workbook or not self.sheet:
            try:
                # Try to reconnect
                connection_status = self.connect()
                if "error" in connection_status.get("status", ""):
                    return {
                        "status": "error",
                        "error": f"Not connected to Excel: {connection_status.get('message', '')}"
                    }
            except Exception as e:
                return {
                    "status": "error",
                    "error": f"Excel connection error: {str(e)}"
                }
        
        command_type = command.get("type")
        
        try:
            # Ensure screen updating is on for this command execution
            prev_screen_updating = self.excel_app.screen_updating
            self.excel_app.screen_updating = True
            
            # Execute the appropriate command
            if command_type == "insert_formula":
                result = self._insert_formula(command)
            elif command_type == "format_range":
                result = self._format_range(command)
            elif command_type == "create_chart":
                result = self._create_chart(command)
            elif command_type == "get_data":
                result = self._get_data(command)
            elif command_type == "insert_data":
                result = self._insert_data(command)
            elif command_type == "sort_data":
                result = self._sort_data(command)
            elif command_type == "filter_data":
                result = self._filter_data(command)
            else:
                result = {"error": f"Unknown command type: {command_type}", "status": "error"}
            
            # Restore previous screen updating setting
            self.excel_app.screen_updating = prev_screen_updating
            
            return result
        except Exception as e:
            # Capture any exceptions during command execution
            return {
                "status": "error",
                "error": f"Error executing {command_type} command: {str(e)}"
            }
    
    def _insert_formula(self, command: Dict[str, Any]) -> Dict[str, Any]:
        """
        Insert a formula into a cell or range.
        
        Args:
            command: A dictionary with cell address and formula
            
        Returns:
            dict: The result of the operation
        """
        try:
            cell_address = command.get("cell")
            formula = command.get("formula")
            
            if not cell_address or not formula:
                return {"error": "Cell address and formula are required", "status": "error"}
            
            # Insert the formula using xlwings
            self.sheet.range(cell_address).formula = formula
            
            return {
                "status": "success",
                "message": f"Formula inserted at {cell_address}"
            }
        except Exception as e:
            return {"error": str(e), "status": "error"}
    
    def _format_range(self, command: Dict[str, Any]) -> Dict[str, Any]:
        """
        Format a range of cells.
        
        Args:
            command: A dictionary with range and formatting options
            
        Returns:
            dict: The result of the operation
        """
        try:
            range_address = command.get("range")
            formatting = command.get("formatting", {})
            
            if not range_address:
                return {"error": "Range address is required", "status": "error"}
            
            # Apply formatting using xlwings
            range_obj = self.sheet.range(range_address)
            
            if "number_format" in formatting:
                range_obj.number_format = formatting["number_format"]
            
            if "font_bold" in formatting:
                range_obj.api.Font.Bold = formatting["font_bold"]
            
            if "font_color" in formatting:
                range_obj.api.Font.Color = formatting["font_color"]
            
            if "fill_color" in formatting:
                range_obj.color = formatting["fill_color"]
            
            return {
                "status": "success",
                "message": f"Formatting applied to {range_address}"
            }
        except Exception as e:
            return {"error": str(e), "status": "error"}
    
    def _create_chart(self, command: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create a chart from data in the specified range.
        
        Args:
            command: A dictionary with chart options
            
        Returns:
            dict: The result of the operation
        """
        try:
            data_range = command.get("data_range")
            chart_type = command.get("chart_type", "column")
            chart_position = command.get("position", "A10")
            
            if not data_range:
                return {"error": "Data range is required", "status": "error"}
            
            # Create chart using xlwings
            chart = self.sheet.charts.add(left=100, top=100, width=375, height=225)
            chart.set_source_data(self.sheet.range(data_range))
            
            # Set chart type
            chart_type_map = {
                "column": "column_clustered",
                "bar": "bar_clustered", 
                "line": "line",
                "pie": "pie",
                "scatter": "xy_scatter",
                "area": "area"
            }
            
            # Set the chart type using xlwings' friendly names
            chart.chart_type = chart_type_map.get(chart_type, "column_clustered")
            
            return {
                "status": "success",
                "message": f"Chart created from data range {data_range}"
            }
        except Exception as e:
            return {"error": str(e), "status": "error"}
    
    def _get_data(self, command: Dict[str, Any]) -> Dict[str, Any]:
        """
        Get data from a range in the Excel sheet.
        
        Args:
            command: A dictionary with range specification
            
        Returns:
            dict: The data from the specified range
        """
        try:
            range_address = command.get("range")
            
            if not range_address:
                return {"error": "Range address is required", "status": "error"}
            
            # Get the data using xlwings
            data = self.sheet.range(range_address).value
            
            # Ensure we have a 2D list
            if not isinstance(data, list):
                data = [[data]]
            elif data and not isinstance(data[0], list):
                data = [data]
            
            return {
                "status": "success",
                "data": data
            }
        except Exception as e:
            return {"error": str(e), "status": "error"}
    
    def _insert_data(self, command: Dict[str, Any]) -> Dict[str, Any]:
        """
        Insert data into cells.
        
        Args:
            command: A dictionary with range and data to insert
            
        Returns:
            dict: The result of the operation
        """
        try:
            range_address = command.get("range")
            data = command.get("data")
            
            if not range_address:
                # Try to get from parameters if using pattern matching
                params = command.get("parameters", [])
                if len(params) >= 2:
                    data = params[0]
                    range_address = params[1]
            
            if not range_address or data is None:
                return {"error": "Range address and data are required", "status": "error"}
            
            # Insert the data using xlwings
            self.sheet.range(range_address).value = data
            
            return {
                "status": "success",
                "message": f"Data inserted at {range_address}"
            }
        except Exception as e:
            return {"error": str(e), "status": "error"}
    
    def _sort_data(self, command: Dict[str, Any]) -> Dict[str, Any]:
        """
        Sort data in a range.
        
        Args:
            command: A dictionary with range and sort key information
            
        Returns:
            dict: The result of the operation
        """
        try:
            range_address = command.get("range")
            sort_column = command.get("sort_column")
            sort_order = command.get("sort_order", "ascending")
            
            if not range_address or sort_column is None:
                return {"error": "Range address and sort_column are required", "status": "error"}
            
            # Convert sort order to Excel constant
            order_const = 1 if sort_order.lower() == "ascending" else 2  # xlAscending=1, xlDescending=2
            
            # Get active sheet
            xl_app = self.excel_app
            sheet = xl_app.ActiveSheet
            
            # Apply the sort
            sheet.Range(range_address).Sort(
                Key1=sheet.Range(range_address).Columns(sort_column),
                Order1=order_const,
                Header=1  # xlYes (assumes header row)
            )
            
            return {
                "status": "success", 
                "message": f"Sorted data in {range_address} by column {sort_column} in {sort_order} order",
                "range": range_address,
                "sort_column": sort_column,
                "sort_order": sort_order
            }
        except Exception as e:
            return {"error": str(e), "status": "error"}
    
    def _filter_data(self, command: Dict[str, Any]) -> Dict[str, Any]:
        """
        Filter data in a range.
        
        Args:
            command: A dictionary with range and filter criteria
            
        Returns:
            dict: The result of the operation
        """
        try:
            range_address = command.get("range")
            filter_column = command.get("filter_column")
            criteria = command.get("criteria")
            
            if not range_address:
                # Try to get from parameters if using pattern matching
                params = command.get("parameters", [])
                if len(params) >= 2:
                    range_address = params[0]
                    criteria_text = params[1]
                    # Try to parse criteria from text
                    criteria_parts = criteria_text.split()
                    if len(criteria_parts) >= 3:
                        filter_column = criteria_parts[0]
                        criteria = " ".join(criteria_parts[2:])
            
            if not range_address or not filter_column or not criteria:
                return {"error": "Range address, filter column, and criteria are required", "status": "error"}
            
            # Get the range using xlwings
            range_obj = self.sheet.range(range_address)
            
            # Convert filter_column to index if it's a letter
            column_index = None
            if isinstance(filter_column, str) and filter_column.isalpha():
                # Convert column letter to 0-based index (A=0, B=1, etc.)
                column_index = 0
                for char in filter_column.upper():
                    column_index = column_index * 26 + (ord(char) - ord('A'))
            elif isinstance(filter_column, str) and filter_column.isdigit():
                column_index = int(filter_column) - 1  # Convert to 0-based index
            else:
                # Try to find column by name in header row
                header_row = range_obj[0, :].value
                if isinstance(header_row, list):
                    for i, header in enumerate(header_row):
                        if str(header).lower() == filter_column.lower():
                            column_index = i
                            break
            
            if column_index is None:
                return {"error": f"Could not find column {filter_column}", "status": "error"}
            
            # Apply autofilter through the Excel API since xlwings doesn't have direct filter methods
            # First, ensure AutoFilter is applied
            if not range_obj.api.AutoFilter():
                range_obj.api.AutoFilter()
            
            # Then apply the specific filter criteria
            # Field is 1-indexed column in the range
            range_obj.api.AutoFilter(Field=column_index + 1, Criteria1=criteria)
            
            return {
                "status": "success",
                "message": f"Data in {range_address} filtered by {filter_column} with criteria: {criteria}"
            }
        except Exception as e:
            return {"error": str(e), "status": "error"}
    
    def get_health_status(self):
        """Get the health status of the Excel AI Agent"""
        try:
            excel_status = "Not connected"
            workbook_name = "None"
            sheet_name = "None"
            
            if self.excel_app:
                try:
                    # Don't use excel_app.name as it doesn't exist
                    excel_status = "Connected"
                    
                    if self.workbook:
                        try:
                            workbook_name = self.workbook.Name
                        except:
                            workbook_name = "Error getting workbook name"
                    
                    if self.sheet:
                        try:
                            sheet_name = self.sheet.Name
                        except:
                            sheet_name = "Error getting sheet name"
                except Exception as e:
                    excel_status = f"Error: {str(e)}"
            
            return {
                "status": "OK",
                "excel_connected": excel_status,
                "workbook": workbook_name,
                "sheet": sheet_name,
                "version": self.version
            }
        except Exception as e:
            return {
                "status": "Error",
                "message": f"Error getting health status: {str(e)}",
                "version": self.version
            }
    
    def _trigger_autonomous_actions(self, query: str) -> Dict[str, Any]:
        """
        Trigger autonomous actions in Excel based on the user's query.
        
        This method:
        1. Gets Excel context
        2. Sends the query to the AI to get recommended actions
        3. Parses and executes each action
        4. Returns the results
        
        Args:
            query: The user's natural language query
            
        Returns:
            Dict: Results of the autonomous actions
        """
        try:
            # Get Excel context
            context = self._get_excel_context()
            
            # Define system prompt for action generation
            system_prompt = """
            You are an Excel Action Generator. Given a user query and their Excel context, 
            recommend specific actions to perform in Excel. Return a JSON array of actions.
            
            Each action should be a JSON object with:
            - "action_type": One of "insert_formula", "format_range", "create_chart", "get_data", 
                            "insert_data", "sort_data", "filter_data"
            - Parameters specific to each action type:
                * insert_formula: "cell", "formula"
                * format_range: "range", "format_options" (object with optional "number_format", "color", "bold", "italic", "alignment")
                * create_chart: "chart_type", "data_range", "title", "position"
                * get_data: "range"
                * insert_data: "start_cell", "data" (2D array)
                * sort_data: "range", "key_column", "ascending" (boolean)
                * filter_data: "range", "column", "criteria"
            
            Be specific and precise with cell references and formulas.
            Be sure to include all required parameters for each action type.
            Return only valid actions based on the Excel context provided.
            """
            
            # Get AI-suggested actions
            ai_response = self._get_ai_action_response(query, system_prompt, context)
            
            # Parse the AI response to extract actions
            actions = self._parse_ai_actions(ai_response)
            
            if not actions:
                return {
                    "status": "error",
                    "message": "No valid actions found in AI response"
                }
            
            # Execute each action and collect results
            results = []
            for action in actions:
                result = self._execute_excel_action(action)
                results.append(result)
            
            # Prepare a summary of actions performed
            success_count = sum(1 for r in results if r.get("status") == "success")
            failed_count = len(results) - success_count
            
            return {
                "status": "success" if success_count > 0 else "error",
                "message": f"Executed {success_count} actions successfully, {failed_count} failed",
                "actions_attempted": len(actions),
                "actions_succeeded": success_count,
                "actions_failed": failed_count,
                "results": results
            }
            
        except Exception as e:
            logging.error(f"Error triggering autonomous actions: {str(e)}")
            return {
                "status": "error",
                "message": f"Error triggering autonomous actions: {str(e)}"
            }
    
    def _parse_ai_actions(self, ai_response: str) -> List[Dict[str, Any]]:
        """
        Parse the AI response to extract the actions.
        
        Args:
            ai_response: The response from the AI model
            
        Returns:
            List[Dict]: List of action dictionaries
        """
        try:
            # Try to find JSON array in the response
            # Look for content between square brackets
            match = re.search(r'\[\s*{.*}\s*\]', ai_response, re.DOTALL)
            if match:
                json_str = match.group(0)
                actions = json.loads(json_str)
                return actions
            
            # If no array found, try to parse as a single JSON object
            match = re.search(r'{.*}', ai_response, re.DOTALL)
            if match:
                json_str = match.group(0)
                action = json.loads(json_str)
                return [action]
            
            # If still no valid JSON found, return empty list
            return []
            
        except Exception as e:
            logging.error(f"Error parsing AI actions: {str(e)}")
            return []
    
    def _get_ai_action_response(self, query: str, system_prompt: str, context: Dict[str, Any]) -> str:
        """
        Get AI-suggested actions based on the user query and Excel context.
        
        Args:
            query: The user's query
            system_prompt: The system prompt for action generation
            context: The Excel context (selection, sheet data, etc.)
            
        Returns:
            str: The AI's response with suggested actions
        """
        try:
            # Convert context to a formatted string
            context_str = json.dumps(context, indent=2)
            
            # Build the full prompt for the AI
            full_prompt = f"""
            {system_prompt}
            
            USER QUERY: {query}
            
            EXCEL CONTEXT:
            {context_str}
            
            Based on this information, generate a JSON array of specific Excel actions to perform.
            Each action must include all required parameters for its action_type.
            """
            
            # Generate content using Gemini API
            genai.configure(api_key=self.api_key)
            model = genai.GenerativeModel('gemini-pro')
            
            # Add safety settings to avoid harmful content
            safety_settings = {
                generative_models.HarmCategory.HARM_CATEGORY_HARASSMENT: 
                    generative_models.HarmBlockThreshold.BLOCK_NONE,
                generative_models.HarmCategory.HARM_CATEGORY_HATE_SPEECH: 
                    generative_models.HarmBlockThreshold.BLOCK_NONE,
                generative_models.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: 
                    generative_models.HarmBlockThreshold.BLOCK_NONE,
                generative_models.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: 
                    generative_models.HarmBlockThreshold.BLOCK_NONE,
            }
            
            response = model.generate_content(
                full_prompt,
                generation_config=genai.types.GenerationConfig(
                    temperature=0.2,
                    top_p=0.95,
                    top_k=40,
                    max_output_tokens=2048,
                ),
                safety_settings=safety_settings
            )
            
            # Extract the response text
            if response.candidates and response.candidates[0].content.parts:
                return response.candidates[0].content.parts[0].text
            else:
                logging.error("Empty response from Gemini API")
                return "[]"  # Return empty array as string
                
        except Exception as e:
            logging.error(f"Error getting AI action response: {str(e)}")
            return "[]"  # Return empty array as string in case of error
    
    def _execute_excel_action(self, action: Dict[str, Any]) -> Dict[str, Any]:
        """
        Execute a single Excel action based on the action dictionary.
        
        Args:
            action: Dictionary containing action_type and parameters for the action
            
        Returns:
            Dict with success status, result message, and any output data
        """
        try:
            # Validate action structure
            if not isinstance(action, dict) or 'action_type' not in action:
                return {"success": False, "message": "Invalid action format: missing action_type"}
            
            action_type = action.get('action_type', '').lower()
            
            # Execute appropriate action based on action_type
            if action_type == 'insert_formula':
                return self._action_insert_formula(action)
            elif action_type == 'format_range':
                return self._action_format_range(action)
            elif action_type == 'create_chart':
                return self._action_create_chart(action)
            elif action_type == 'filter_data':
                return self._action_filter_data(action)
            elif action_type == 'sort_data':
                return self._action_sort_data(action)
            elif action_type == 'insert_text':
                return self._action_insert_text(action)
            elif action_type == 'create_pivot_table':
                return self._action_create_pivot_table(action)
            elif action_type == 'create_table':
                return self._action_create_table(action)
            else:
                return {"success": False, "message": f"Unsupported action type: {action_type}"}
                
        except Exception as e:
            logging.error(f"Error executing Excel action: {str(e)}")
            return {"success": False, "message": f"Error: {str(e)}"}
    
    def _action_insert_formula(self, action: Dict[str, Any]) -> Dict[str, Any]:
        """Insert a formula into specified cells"""
        try:
            # Extract required parameters
            cell_range = action.get('range')
            formula = action.get('formula')
            
            # Validate parameters
            if not cell_range or not formula:
                return {"success": False, "message": "Missing required parameters: range and formula"}
            
            # Get active sheet
            xl_app = self.excel_app
            sheet = xl_app.ActiveSheet
            
            # Insert formula
            sheet.Range(cell_range).Formula = formula
            
            return {
                "success": True, 
                "message": f"Formula '{formula}' inserted into {cell_range}",
                "cell_range": cell_range,
                "formula": formula
            }
        except Exception as e:
            return {"success": False, "message": f"Error inserting formula: {str(e)}"}
    
    def _action_format_range(self, action: Dict[str, Any]) -> Dict[str, Any]:
        """Apply formatting to a specified range of cells"""
        try:
            # Extract required parameters
            cell_range = action.get('range')
            format_type = action.get('format_type', '')
            
            if not cell_range or not format_type:
                return {"success": False, "message": "Missing required parameters: range and format_type"}
            
            # Get active sheet
            xl_app = self.excel_app
            sheet = xl_app.ActiveSheet
            target_range = sheet.Range(cell_range)
            
            # Apply formatting based on format_type
            if format_type.lower() == 'bold':
                target_range.Font.Bold = True
            elif format_type.lower() == 'italic':
                target_range.Font.Italic = True
            elif format_type.lower() == 'underline':
                target_range.Font.Underline = True
            elif format_type.lower() == 'number':
                number_format = action.get('number_format', '0.00')
                target_range.NumberFormat = number_format
            elif format_type.lower() == 'color':
                color_index = action.get('color_index', 3)  # Default to red if not specified
                target_range.Interior.ColorIndex = color_index
            elif format_type.lower() == 'font_color':
                color_index = action.get('color_index', 3)  # Default to red if not specified
                target_range.Font.ColorIndex = color_index
            else:
                return {"success": False, "message": f"Unsupported format type: {format_type}"}
            
            return {
                "success": True, 
                "message": f"Applied {format_type} formatting to {cell_range}",
                "cell_range": cell_range,
                "format_type": format_type
            }
        except Exception as e:
            return {"success": False, "message": f"Error applying formatting: {str(e)}"}
    
    def _action_create_chart(self, action: Dict[str, Any]) -> Dict[str, Any]:
        """Create a chart from data in specified range"""
        try:
            # Extract required parameters
            data_range = action.get('data_range')
            chart_type = action.get('chart_type', 'xlColumnClustered')
            chart_position = action.get('position', 'A10')
            chart_title = action.get('title', 'Chart')
            
            if not data_range:
                return {"success": False, "message": "Missing required parameter: data_range"}
            
            # Map string chart types to Excel constants
            chart_type_map = {
                'column': -4100,  # xlColumnClustered
                'line': 4,        # xlLine
                'pie': 5,         # xlPie
                'bar': 57,        # xlBarClustered
                'area': 76,       # xlAreaStacked
                'scatter': -4169  # xlXYScatter
            }
            
            # Get Excel constants from string
            chart_type_const = chart_type_map.get(chart_type.lower(), -4100)  # Default to column chart
            
            # Get active sheet
            xl_app = self.excel_app
            sheet = xl_app.ActiveSheet
            
            # Create the chart
            charts = sheet.ChartObjects()
            chart = charts.Add(
                Left=sheet.Range(chart_position).Left, 
                Top=sheet.Range(chart_position).Top,
                Width=300, 
                Height=200
            )
            
            # Set chart properties
            chart_obj = chart.Chart
            chart_obj.SetSourceData(sheet.Range(data_range))
            chart_obj.ChartType = chart_type_const
            chart_obj.HasTitle = True
            chart_obj.ChartTitle.Text = chart_title
            
            return {
                "success": True, 
                "message": f"Created {chart_type} chart from data in {data_range}",
                "data_range": data_range,
                "chart_type": chart_type,
                "position": chart_position
            }
        except Exception as e:
            return {"success": False, "message": f"Error creating chart: {str(e)}"}
    
    def _action_filter_data(self, action: Dict[str, Any]) -> Dict[str, Any]:
        """Apply a filter to a data range"""
        try:
            # Extract required parameters
            data_range = action.get('range')
            column_index = action.get('column_index')
            filter_criteria = action.get('criteria')
            
            if not data_range or column_index is None or not filter_criteria:
                return {"success": False, "message": "Missing required parameters: range, column_index, and criteria"}
            
            # Get active sheet
            xl_app = self.excel_app
            sheet = xl_app.ActiveSheet
            
            # Create or use existing filter
            data_range_obj = sheet.Range(data_range)
            if not sheet.AutoFilterMode:
                data_range_obj.AutoFilter()
            
            # Apply the filter
            # Excel's column indices are 1-based
            column_index = int(column_index)
            sheet.Range(data_range).AutoFilter(Field=column_index, Criteria1=filter_criteria)
            
            return {
                "success": True, 
                "message": f"Applied filter on column {column_index} with criteria '{filter_criteria}'",
                "range": data_range,
                "column_index": column_index,
                "criteria": filter_criteria
            }
        except Exception as e:
            return {"success": False, "message": f"Error applying filter: {str(e)}"}
    
    def _action_sort_data(self, action: Dict[str, Any]) -> Dict[str, Any]:
        """Sort a range of data"""
        try:
            # Extract required parameters
            data_range = action.get('range')
            sort_column = action.get('sort_column')
            sort_order = action.get('sort_order', 'ascending')
            
            if not data_range or sort_column is None:
                return {"success": False, "message": "Missing required parameters: range and sort_column"}
            
            # Convert sort order to Excel constant
            order_const = 1 if sort_order.lower() == 'ascending' else 2  # xlAscending=1, xlDescending=2
            
            # Get active sheet
            xl_app = self.excel_app
            sheet = xl_app.ActiveSheet
            
            # Apply the sort
            sheet.Range(data_range).Sort(
                Key1=sheet.Range(data_range).Columns(sort_column),
                Order1=order_const,
                Header=1  # xlYes (assumes header row)
            )
            
            return {
                "success": True, 
                "message": f"Sorted data in {data_range} by column {sort_column} in {sort_order} order",
                "range": data_range,
                "sort_column": sort_column,
                "sort_order": sort_order
            }
        except Exception as e:
            return {"success": False, "message": f"Error sorting data: {str(e)}"}
    
    def _execute_sort_data(self, action: Dict[str, Any]) -> Dict[str, Any]:
        """
        Sort data in Excel.
        
        Args:
            action: Dictionary with 'range', 'key_column', and optionally 'ascending' keys
            
        Returns:
            Dict: Result of the sort operation
        """
        try:
            range_address = action.get("range", "")
            key_column = action.get("key_column", 1)  # 1-based column index
            ascending = action.get("ascending", True)
            
            if not range_address:
                return {"status": "error", "message": "Missing range parameter"}
            
            sheet = xw.books.active.sheets.active
            range_obj = sheet.range(range_address)
            
            # Convert key_column to 0-based index if it's a number
            if isinstance(key_column, int):
                key_column_index = key_column - 1
            else:
                # If key_column is a letter (like 'A', 'B', etc.)
                key_column_index = ord(key_column.upper()) - ord('A')
            
            # Use Excel's native sorting
            sort_order = 1 if ascending else 2  # 1 for ascending, 2 for descending
            range_obj.api.Sort(
                Key1=range_obj.api.Cells(1, key_column_index + 1),
                Order1=sort_order,
                Header=1,  # xlYes - Assume headers are present
                OrderCustom=1,
                MatchCase=False,
                Orientation=1,  # xlTopToBottom
                SortMethod=1  # xlPinYin
            )
            
            return {
                "status": "success", 
                "message": f"Data in range {range_address} sorted by column {key_column} ({'ascending' if ascending else 'descending'})",
                "sort_details": {
                    "range": range_address,
                    "key_column": key_column,
                    "ascending": ascending
                }
            }
        except Exception as e:
            return {"status": "error", "message": f"Error sorting data: {str(e)}"}
    
    def _execute_filter_data(self, action: Dict[str, Any]) -> Dict[str, Any]:
        """
        Apply a filter to data in Excel.
        
        Args:
            action: Dictionary with 'range', 'column', and 'criteria' keys
            
        Returns:
            Dict: Result of the filter operation
        """
        try:
            range_address = action.get("range", "")
            column = action.get("column", "")  # Can be letter or number
            criteria = action.get("criteria", "")
            
            if not range_address or not column or not criteria:
                return {"status": "error", "message": "Missing required parameters for filtering"}
            
            sheet = xw.books.active.sheets.active
            range_obj = sheet.range(range_address)
            
            # Convert column to index if it's a letter
            if isinstance(column, str) and column.isalpha():
                column_index = sum([(ord(c.upper()) - ord('A') + 1) * (26 ** i) for i, c in enumerate(reversed(column.upper()))])
            else:
                # If column is a number, use it directly (convert to int if needed)
                column_index = int(column)
            
            # Make sure autofilter is enabled for the range
            if not range_obj.api.AutoFilter():
                range_obj.api.AutoFilter()
            
            # Apply the filter
            range_obj.api.AutoFilter(Field=column_index, Criteria1=criteria)
            
            return {
                "status": "success", 
                "message": f"Filter applied to column {column} with criteria '{criteria}'",
                "filter_details": {
                    "range": range_address,
                    "column": column,
                    "criteria": criteria
                }
            }
        except Exception as e:
            return {"status": "error", "message": f"Error applying filter: {str(e)}"}

    def _is_excel_connected(self) -> bool:
        """
        Check if Excel is currently connected.
        
        Returns:
            bool: True if Excel is connected, False otherwise
        """
        try:
            if not self.excel_app:
                return False
                
            # Try to access a property to verify the connection
            _ = self.excel_app.ActiveWorkbook.Name
            return True
        except:
            return False 