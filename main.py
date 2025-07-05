import sys
import unohelper
import officehelper
import json
import urllib.request
import urllib.parse
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
import uno
import os 
import logging
import re

# Import LiteLLM (to be vendored in lib/)
try:
    sys.path.append(os.path.join(os.path.dirname(__file__), 'lib'))
    from litellm import completion
except ImportError:
    # Fallback or error handling if LiteLLM is not available
    def completion(*args, **kwargs):
        raise Exception("LiteLLM library not found. Please ensure it is installed in the lib/ directory.")

from com.sun.star.beans import PropertyValue
from com.sun.star.container import XNamed


def log_to_file(message):
    # Get the user's home directory
    home_directory = os.path.expanduser('~')
    
    # Define the log file path
    log_file_path = os.path.join(home_directory, 'log.txt')
    
    # Set up logging configuration
    logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(message)s')
    
    # Log the input message
    logging.info(message)


# The MainJob is a UNO component derived from unohelper.Base class
# and also the XJobExecutor, the implemented interface
class MainJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx
        # handling different situations (inside LibreOffice or other process)
        try:
            self.sm = ctx.getServiceManager()
            self.desktop = XSCRIPTCONTEXT.getDesktop()
            self.document = XSCRIPTCONTEXT.getDocument()
        except NameError:
            self.sm = ctx.ServiceManager
            self.desktop = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)
    

    def call_completion(self, messages, max_tokens=None, endpoint=None, model_name=None, provider=None, api_key=None):
        """Make a completion call with standardized parameters"""
        try:
            # Construct model string based on provider if provided
            full_model = model_name or self.get_config("model", "")
            if provider and full_model:
                full_model = f"{provider}/{full_model}"
            elif not full_model:
                # Default to a generic OpenAI-compatible model with explicit provider
                full_model = "openai/gpt-3.5-turbo"
                if provider:
                    full_model = f"{provider}/gpt-3.5-turbo"

            kwargs = {
                "model": full_model,
                "messages": messages,
                "max_tokens": max_tokens or self.get_config("extend_selection_max_tokens", 70)
            }
            # Only include api_base if endpoint is not empty
            endpoint = endpoint or self.get_config("endpoint", "http://127.0.0.1:5000")
            if endpoint:
                kwargs["api_base"] = endpoint
            api_key = api_key or self.get_config("api_key", "")
            if api_key:
                kwargs["api_key"] = api_key

            return completion(**kwargs)
        except Exception as e:
            raise Exception(f"Completion error: {str(e)}")

    def get_config(self, key, default):
  
        name_file ="localwriter.json"
        #path_settings = create_instance('com.sun.star.util.PathSettings')
        
        
        path_settings = self.sm.createInstanceWithContext('com.sun.star.util.PathSettings', self.ctx)

        user_config_path = getattr(path_settings, "UserConfig")

        if user_config_path.startswith('file://'):
            user_config_path = str(uno.fileUrlToSystemPath(user_config_path))
        
        # Ensure the path ends with the filename
        config_file_path = os.path.join(user_config_path, name_file)

        # Check if the file exists
        if not os.path.exists(config_file_path):
            return default

        # Try to load the JSON content from the file
        try:
            with open(config_file_path, 'r') as file:
                config_data = json.load(file)
        except (IOError, json.JSONDecodeError):
            return default

        # Return the value corresponding to the key, or the default value if the key is not found
        return config_data.get(key, default)

    def set_config(self, key, value):
        name_file = "localwriter.json"
        
        path_settings = self.sm.createInstanceWithContext('com.sun.star.util.PathSettings', self.ctx)
        user_config_path = getattr(path_settings, "UserConfig")

        if user_config_path.startswith('file://'):
            user_config_path = str(uno.fileUrlToSystemPath(user_config_path))

        # Ensure the path ends with the filename
        config_file_path = os.path.join(user_config_path, name_file)

        # Load existing configuration if the file exists
        if os.path.exists(config_file_path):
            try:
                with open(config_file_path, 'r') as file:
                    config_data = json.load(file)
            except (IOError, json.JSONDecodeError):
                config_data = {}
        else:
            config_data = {}

        # Update the configuration with the new key-value pair
        config_data[key] = value

        # Write the updated configuration back to the file
        try:
            with open(config_file_path, 'w') as file:
                json.dump(config_data, file, indent=4)
        except IOError as e:
            # Handle potential IO errors (optional)
            print(f"Error writing to {config_file_path}: {e}")


    #retrieved from https://wiki.documentfoundation.org/Macros/General/IO_to_Screen
    #License: Creative Commons Attribution-ShareAlike 3.0 Unported License,
    #License: The Document Foundation  https://creativecommons.org/licenses/by-sa/3.0/
    #begin sharealike section 
    def input_box(self,message, title="", default="", x=None, y=None):
        """ Shows dialog with input box.
            @param message message to show on the dialog
            @param title window title
            @param default default value
            @param x optional dialog position in twips
            @param y optional dialog position in twips
            @return string if OK button pushed, otherwise zero length string
        """
        WIDTH = 600
        HORI_MARGIN = VERT_MARGIN = 8
        BUTTON_WIDTH = 100
        BUTTON_HEIGHT = 26
        HORI_SEP = VERT_SEP = 8
        LABEL_HEIGHT = BUTTON_HEIGHT * 2 + 5
        EDIT_HEIGHT = 24
        HEIGHT = VERT_MARGIN * 2 + LABEL_HEIGHT + VERT_SEP + EDIT_HEIGHT
        import uno
        from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
        from com.sun.star.awt.PushButtonType import OK, CANCEL
        from com.sun.star.util.MeasureUnit import TWIP
        ctx = uno.getComponentContext()
        def create(name):
            return ctx.getServiceManager().createInstanceWithContext(name, ctx)
        dialog = create("com.sun.star.awt.UnoControlDialog")
        dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
        dialog.setModel(dialog_model)
        dialog.setVisible(False)
        dialog.setTitle(title)
        dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)
        def add(name, type, x_, y_, width_, height_, props):
            model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + type + "Model")
            dialog_model.insertByName(name, model)
            control = dialog.getControl(name)
            control.setPosSize(x_, y_, width_, height_, POSSIZE)
            for key, value in props.items():
                setattr(model, key, value)
        label_width = WIDTH - BUTTON_WIDTH - HORI_SEP - HORI_MARGIN * 2
        add("label", "FixedText", HORI_MARGIN, VERT_MARGIN, label_width, LABEL_HEIGHT, 
            {"Label": str(message), "NoLabel": True})
        add("btn_ok", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN, 
                BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
        add("edit", "Edit", HORI_MARGIN, LABEL_HEIGHT + VERT_MARGIN + VERT_SEP, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(default)})
        frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
        window = frame.getContainerWindow() if frame else None
        dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
        if not x is None and not y is None:
            ps = dialog.convertSizeToPixel(uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP)
            _x, _y = ps.Width, ps.Height
        elif window:
            ps = window.getPosSize()
            _x = ps.Width / 2 - WIDTH / 2
            _y = ps.Height / 2 - HEIGHT / 2
        dialog.setPosSize(_x, _y, 0, 0, POS)
        edit = dialog.getControl("edit")
        edit.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(default))))
        edit.setFocus()
        ret = edit.getModel().Text if dialog.execute() else ""
        dialog.dispose()
        return ret

    def settings_box(self, title="", x=None, y=None):
        """ Shows dialog with input box for settings.
            @param title window title
            @param x optional dialog position in twips
            @param y optional dialog position in twips
            @return dictionary of settings if OK button pushed, otherwise empty dictionary
        """
        WIDTH = 600
        HORI_MARGIN = VERT_MARGIN = 8
        BUTTON_WIDTH = 100
        BUTTON_HEIGHT = 26
        HORI_SEP = 8
        VERT_SEP = 4
        LABEL_HEIGHT = BUTTON_HEIGHT + 5
        EDIT_HEIGHT = 24
        HEIGHT = VERT_MARGIN * 8 + LABEL_HEIGHT * 9 + VERT_SEP * 10 + EDIT_HEIGHT * 8 + 350
        import uno
        from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
        from com.sun.star.awt.PushButtonType import OK, CANCEL
        from com.sun.star.util.MeasureUnit import TWIP
        ctx = uno.getComponentContext()
        def create(name):
            return ctx.getServiceManager().createInstanceWithContext(name, ctx)
        dialog = create("com.sun.star.awt.UnoControlDialog")
        dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
        dialog.setModel(dialog_model)
        dialog.setVisible(False)
        dialog.setTitle(title)
        dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)
        def add(name, type, x_, y_, width_, height_, props):
            model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + type + "Model")
            dialog_model.insertByName(name, model)
            control = dialog.getControl(name)
            control.setPosSize(x_, y_, width_, height_, POSSIZE)
            for key, value in props.items():
                setattr(model, key, value)
        label_width = WIDTH - BUTTON_WIDTH - HORI_SEP - HORI_MARGIN * 2
        add("label_endpoint", "FixedText", HORI_MARGIN, VERT_MARGIN, label_width, LABEL_HEIGHT, 
            {"Label": "Endpoint URL/Port (Local or Proxy, e.g., https://api.x.ai/v1/chat/completions for Grok):", "NoLabel": True})
        add("btn_ok", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN, 
                BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
        add("edit_endpoint", "Edit", HORI_MARGIN, LABEL_HEIGHT,
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("endpoint", "http://127.0.0.1:5000"))})
        
        add("label_model", "FixedText", HORI_MARGIN, LABEL_HEIGHT + VERT_MARGIN + VERT_SEP + EDIT_HEIGHT, label_width, LABEL_HEIGHT, 
            {"Label": "Model (Required by Ollama):", "NoLabel": True})
        add("edit_model", "Edit", HORI_MARGIN, LABEL_HEIGHT*2 + VERT_MARGIN + VERT_SEP*2 + EDIT_HEIGHT, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("model", ""))})
        
        add("label_provider", "FixedText", HORI_MARGIN, LABEL_HEIGHT*3 + VERT_MARGIN + VERT_SEP*3 + EDIT_HEIGHT*2, label_width, LABEL_HEIGHT, 
            {"Label": "Provider (e.g., openai, ollama, anthropic):", "NoLabel": True})
        add("edit_provider", "Edit", HORI_MARGIN, LABEL_HEIGHT*4 + VERT_MARGIN + VERT_SEP*4 + EDIT_HEIGHT*2, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("provider", ""))})
        
        add("label_api_key", "FixedText", HORI_MARGIN, LABEL_HEIGHT*5 + VERT_MARGIN + VERT_SEP*5 + EDIT_HEIGHT*3, label_width, LABEL_HEIGHT, 
            {"Label": "API Key (Optional, use with caution):", "NoLabel": True})
        add("edit_api_key", "Edit", HORI_MARGIN, LABEL_HEIGHT*6 + VERT_MARGIN + VERT_SEP*6 + EDIT_HEIGHT*3, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("api_key", "")), "EchoChar": ord('*')})
        
        add("label_extend_selection_max_tokens", "FixedText", HORI_MARGIN, LABEL_HEIGHT*7 + VERT_MARGIN + VERT_SEP*7 + EDIT_HEIGHT*4, label_width, LABEL_HEIGHT, 
            {"Label": "Extend Selection Max Tokens:", "NoLabel": True})
        add("edit_extend_selection_max_tokens", "Edit", HORI_MARGIN, LABEL_HEIGHT*8 + VERT_MARGIN + VERT_SEP*8 + EDIT_HEIGHT*4, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("extend_selection_max_tokens", "70"))})
        
        add("label_extend_selection_system_prompt", "FixedText", HORI_MARGIN, LABEL_HEIGHT*9 + VERT_MARGIN + VERT_SEP*9 + EDIT_HEIGHT*5, label_width, LABEL_HEIGHT, 
            {"Label": "Extend Selection System Prompt:", "NoLabel": True})
        add("edit_extend_selection_system_prompt", "Edit", HORI_MARGIN, LABEL_HEIGHT*10 + VERT_MARGIN + VERT_SEP*10 + EDIT_HEIGHT*5, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("extend_selection_system_prompt", ""))})

        add("label_edit_selection_max_new_tokens", "FixedText", HORI_MARGIN, LABEL_HEIGHT*11 + VERT_MARGIN + VERT_SEP*11 + EDIT_HEIGHT*6, label_width, LABEL_HEIGHT, 
            {"Label": "Edit Selection Max New Tokens:", "NoLabel": True})
        add("edit_edit_selection_max_new_tokens", "Edit", HORI_MARGIN, LABEL_HEIGHT*12 + VERT_MARGIN + VERT_SEP*12 + EDIT_HEIGHT*6, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("edit_selection_max_new_tokens", "0"))})

        add("label_edit_selection_system_prompt", "FixedText", HORI_MARGIN, LABEL_HEIGHT*13 + VERT_MARGIN + VERT_SEP*13 + EDIT_HEIGHT*7, label_width, LABEL_HEIGHT, 
            {"Label": "Edit Selection System Prompt:", "NoLabel": True})
        add("edit_edit_selection_system_prompt", "Edit", HORI_MARGIN, LABEL_HEIGHT*14 + VERT_MARGIN + VERT_SEP*14 + EDIT_HEIGHT*7, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("edit_selection_system_prompt", ""))})

        add("btn_test", "Button", HORI_MARGIN, LABEL_HEIGHT*15 + VERT_MARGIN + VERT_SEP*15 + EDIT_HEIGHT*8, 
                BUTTON_WIDTH, BUTTON_HEIGHT, {"Label": "Test Connection"})
        add("test_result", "Edit", HORI_MARGIN + BUTTON_WIDTH + HORI_SEP, LABEL_HEIGHT*15 + VERT_MARGIN + VERT_SEP*15 + EDIT_HEIGHT*8, 
                label_width - BUTTON_WIDTH - HORI_SEP, EDIT_HEIGHT * 5, {"Text": "Test result will appear here", "MultiLine": True, "ReadOnly": True})

        frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
        window = frame.getContainerWindow() if frame else None
        dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
        if not x is None and not y is None:
            ps = dialog.convertSizeToPixel(uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP)
            _x, _y = ps.Width, ps.Height
        elif window:
            ps = window.getPosSize()
            _x = ps.Width / 2 - WIDTH / 2
            _y = ps.Height / 2 - HEIGHT / 2
        dialog.setPosSize(_x, _y, 0, 0, POS)
        
        edit_endpoint = dialog.getControl("edit_endpoint")
        edit_endpoint.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("endpoint", "http://127.0.0.1:5000")))))
        
        edit_model = dialog.getControl("edit_model")
        edit_model.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("model", "")))))
        
        edit_provider = dialog.getControl("edit_provider")
        edit_provider.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("provider", "")))))
        
        edit_api_key = dialog.getControl("edit_api_key")
        edit_api_key.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("api_key", "")))))
        
        edit_extend_selection_max_tokens = dialog.getControl("edit_extend_selection_max_tokens")
        edit_extend_selection_max_tokens.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("extend_selection_max_tokens", "70")))))
        
        edit_extend_selection_system_prompt = dialog.getControl("edit_extend_selection_system_prompt")
        edit_extend_selection_system_prompt.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("extend_selection_system_prompt", "")))))
        
        edit_edit_selection_max_new_tokens = dialog.getControl("edit_edit_selection_max_new_tokens")
        edit_edit_selection_max_new_tokens.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("edit_selection_max_new_tokens", "0")))))
        
        edit_edit_selection_system_prompt = dialog.getControl("edit_edit_selection_system_prompt")
        edit_edit_selection_system_prompt.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("edit_selection_system_prompt", "")))))
        
        edit_endpoint.setFocus()

        # Add listener for test button using a UNO-compatible ActionListener
        from com.sun.star.awt import XActionListener
        class TestConnectionListener(unohelper.Base, XActionListener):
            def __init__(self, endpoint_ctrl, model_ctrl, provider_ctrl, api_key_ctrl, result_ctrl, main_job):
                self.endpoint_ctrl = endpoint_ctrl
                self.model_ctrl = model_ctrl
                self.provider_ctrl = provider_ctrl
                self.api_key_ctrl = api_key_ctrl
                self.result_ctrl = result_ctrl
                self.main_job = main_job

            def disposing(self, source):
                pass

            def actionPerformed(self, event):
                endpoint = self.endpoint_ctrl.getModel().Text
                model_name = self.model_ctrl.getModel().Text
                provider = self.provider_ctrl.getModel().Text
                api_key = self.api_key_ctrl.getModel().Text
                try:
                    # Turn on debug for this test call
                    import litellm
                    litellm._turn_on_debug()

                    response = self.main_job.call_completion(
                        messages=[{"role": "user", "content": "Hello, are you working?"}],
                        max_tokens=10,
                        endpoint=endpoint,
                        model_name=model_name,
                        provider=provider,
                        api_key=api_key
                    )
                    self.result_ctrl.setText("Success: " + response.choices[0].message.content)
                except Exception as e:
                    self.result_ctrl.setText("Failed: " + str(e))
                    print(f"Test Connection Error: {str(e)}")
                finally:
                    # Turn off debug after the test call
                    litellm._turn_off_debug()

        btn_test = dialog.getControl("btn_test")
        test_listener = TestConnectionListener(edit_endpoint, edit_model, edit_provider, edit_api_key, dialog.getControl("test_result"), self)
        btn_test.addActionListener(test_listener)

        if dialog.execute():
            result = {
                "endpoint": edit_endpoint.getModel().Text,
                "model": edit_model.getModel().Text,
                "provider": edit_provider.getModel().Text,
                "api_key": edit_api_key.getModel().Text,
                "extend_selection_system_prompt": edit_extend_selection_system_prompt.getModel().Text,
                "edit_selection_system_prompt": edit_edit_selection_system_prompt.getModel().Text
            }
            if edit_extend_selection_max_tokens.getModel().Text.isdigit():
                result["extend_selection_max_tokens"] = int(edit_extend_selection_max_tokens.getModel().Text)
            if edit_edit_selection_max_new_tokens.getModel().Text.isdigit():
                result["edit_selection_max_new_tokens"] = int(edit_edit_selection_max_new_tokens.getModel().Text)
        else:
            result = {}

        dialog.dispose()
        return result
    #end sharealike section 

    def trigger(self, args):
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)
        model = desktop.getCurrentComponent()
        #if not hasattr(model, "Text"):
        #    model = self.desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())

        if hasattr(model, "Text"):
            text = model.Text
            selection = model.CurrentController.getSelection()
            text_range = selection.getByIndex(0)

            
            if args == "ExtendSelection":
                # Access the current selection
                #selection = model.CurrentController.getSelection()
                
                if len(text_range.getString()) > 0:
                    # Get the first range of the selection
                    #text_range = selection.getByIndex(0)
                    try:
                        messages = []
                        if self.get_config("extend_selection_system_prompt", "") != "":
                            messages.append({"role": "system", "content": self.get_config("extend_selection_system_prompt", "")})
                        messages.append({"role": "user", "content": text_range.getString()})
                        
                        response = self.call_completion(
                            messages=messages,
                            max_tokens=self.get_config("extend_selection_max_tokens", 70)
                        )

                        # Append completion to selection
                        selected_text = text_range.getString()
                        new_text = selected_text + response.choices[0].message.content
                        text_range.setString(new_text)
                
                    except Exception as e:
                        text_range = selection.getByIndex(0)
                        error_msg = f":error: {str(e)}"
                        print(f"Error in ExtendSelection: {error_msg}")
                        text_range.setString(text_range.getString() + error_msg)

            elif args == "EditSelection":
                # Access the current selection
                try:
                    user_input = self.input_box("Please enter edit instructions!", "Input", "")
                    messages = []
                    if self.get_config("edit_selection_system_prompt", "") != "":
                        messages.append({"role": "system", "content": self.get_config("edit_selection_system_prompt", "")})
                    user_prompt = f"ORIGINAL VERSION:\n{text_range.getString()}\nBelow is an edited version according to the following instructions. There are no comments in the edited version.\nInstructions:\n{user_input}\nEDITED VERSION:"
                    messages.append({"role": "user", "content": user_prompt})

                    response = self.call_completion(
                        messages=messages,
                        max_tokens=len(text_range.getString()) + self.get_config("edit_selection_max_new_tokens", 0)
                    )

                    # Replace selection with completion
                    new_text = response.choices[0].message.content
                    text_range.setString(new_text)

                except Exception as e:
                    text_range = selection.getByIndex(0)
                    error_msg = f":error: {str(e)}"
                    print(f"Error in EditSelection: {error_msg}")
                    text_range.setString(text_range.getString() + error_msg)
            
            elif args == "settings":
                try:
                    result = self.settings_box("Settings")
                                    
                    if "extend_selection_max_tokens" in result:
                        self.set_config("extend_selection_max_tokens", result["extend_selection_max_tokens"])

                    if "extend_selection_system_prompt" in result:
                        self.set_config("extend_selection_system_prompt", result["extend_selection_system_prompt"])

                    if "edit_selection_max_new_tokens" in result:
                        self.set_config("edit_selection_max_new_tokens", result["edit_selection_max_new_tokens"])

                    if "edit_selection_system_prompt" in result:
                        self.set_config("edit_selection_system_prompt", result["edit_selection_system_prompt"])

                    if "endpoint" in result and result["endpoint"].startswith("http"):
                        # Ensure endpoint ends with a single slash to match LiteLLM's behavior
                        endpoint = result["endpoint"].rstrip('/')
                        self.set_config("endpoint", endpoint)

                    if "model" in result:                
                        self.set_config("model", result["model"])
                        
                    if "provider" in result:
                        self.set_config("provider", result["provider"])
                        
                    if "api_key" in result:
                        self.set_config("api_key", result["api_key"])

                except Exception as e:
                    text_range = selection.getByIndex(0)
                    text_range.setString(text_range.getString() + ":error: " + str(e))
        elif hasattr(model, "Sheets"):
            try:
                #text_range.setString(text_range.getString() + ": " + user_input)
                endpoint = self.get_config("endpoint", "http://127.0.0.1:5000")
                model_name = self.get_config("model", "")
                # Get the active sheet
                sheet = model.CurrentController.ActiveSheet
                
                # Get the current selection (which could be a range of cells)
                selection = model.CurrentController.Selection
                

                if args == "EditSelection":
                    user_input= self.input_box("Please enter edit instructions!", "Input", "")


                area = selection.getRangeAddress()
                start_row = area.StartRow
                end_row = area.EndRow
                start_col = area.StartColumn
                end_col = area.EndColumn

                col_range = range(start_col, end_col + 1)
                row_range = range(start_row, end_row + 1)

                for row in row_range:
                    new_values = []
                    for col in col_range:
                        cell = sheet.getCellByPosition(col, row)


                        if args == "ExtendSelection":
                            
                            if len(cell.getString()) > 0:
                                try:
                                    messages = []
                                    if self.get_config("extend_selection_system_prompt", "") != "":
                                        messages.append({"role": "system", "content": self.get_config("extend_selection_system_prompt", "")})
                                    messages.append({"role": "user", "content": cell.getString()})
                                    
                                    response = self.call_completion(
                                        messages=messages,
                                        max_tokens=self.get_config("extend_selection_max_tokens", 70)
                                    )

                                    # Append completion to selection
                                    selected_text = cell.getString()
                                    new_text = selected_text + response.choices[0].message.content
                                    cell.setString(new_text)
                                except Exception as e:
                                    error_msg = f":error: {str(e)}"
                                    print(f"Error in ExtendSelection (Calc): {error_msg}")
                                    cell.setString(cell.getString() + error_msg)
                        elif args == "EditSelection":
                            # Access the current selection
                            try:
                                messages = []
                                if self.get_config("edit_selection_system_prompt", "") != "":
                                    messages.append({"role": "system", "content": self.get_config("edit_selection_system_prompt", "")})
                                user_prompt = f"ORIGINAL VERSION:\n{cell.getString()}\nBelow is an edited version according to the following instructions. Don't waste time thinking, be as fast as you can. There are no comments in the edited version.\nInstructions:\n{user_input}\nEDITED VERSION:"
                                messages.append({"role": "user", "content": user_prompt})

                                response = self.call_completion(
                                    messages=messages,
                                    max_tokens=len(cell.getString()) + self.get_config("edit_selection_max_new_tokens", 0)
                                )

                                # Get previous selected text
                                selected_text = cell.getString()
                                raw_response = response.choices[0].message.content

                                # Action, rather than thought
                                new_text = re.sub(r'<think>.*?</think>', '', raw_response, flags=re.DOTALL)
                                cell.setString(new_text)

                            except Exception as e:
                                error_msg = f":error: {str(e)}"
                                print(f"Error in EditSelection (Calc): {error_msg}")
                                cell.setString(cell.getString() + error_msg)
                        
                        elif args == "settings":
                            try:
                                result = self.settings_box("Settings")
                                                
                                if "extend_selection_max_tokens" in result:
                                    self.set_config("extend_selection_max_tokens", result["extend_selection_max_tokens"])

                                if "extend_selection_system_prompt" in result:
                                    self.set_config("extend_selection_system_prompt", result["extend_selection_system_prompt"])

                                if "edit_selection_max_new_tokens" in result:
                                    self.set_config("edit_selection_max_new_tokens", result["edit_selection_max_new_tokens"])

                                if "edit_selection_system_prompt" in result:
                                    self.set_config("edit_selection_system_prompt", result["edit_selection_system_prompt"])

                                if "endpoint" in result and result["endpoint"].startswith("http"):
                                    # Ensure endpoint ends with a single slash to match LiteLLM's behavior
                                    endpoint = result["endpoint"].rstrip('/')
                                    self.set_config("endpoint", endpoint)

                                if "model" in result:                
                                    self.set_config("model", result["model"])
                                    
                                if "provider" in result:
                                    self.set_config("provider", result["provider"])
                                    
                                if "api_key" in result:
                                    self.set_config("api_key", result["api_key"])

                            except Exception as e:
                                cell.setString(cell.getString() + ":error: " + str(e))
            except Exception as e:
                pass

# Starting from Python IDE
def main():
    try:
        ctx = XSCRIPTCONTEXT
    except NameError:
        ctx = officehelper.bootstrap()
        if ctx is None:
            print("ERROR: Could not bootstrap default Office.")
            sys.exit(1)
    job = MainJob(ctx)
    job.trigger("hello")
# Starting from command line
if __name__ == "__main__":
    main()
# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    MainJob,  # UNO object class
    "org.extension.sample.do",  # implementation name (customize for yourself)
    ("com.sun.star.task.Job",), )  # implemented services (only 1)
# vim: set shiftwidth=4 softtabstop=4 expandtab:
