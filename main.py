import sys
import unohelper
import officehelper
import json
import urllib.request
import urllib.parse
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS, XItemListener
import uno
import os 
import logging
import re

# Import LiteLLM (to be vendored in lib/)
try:
    sys.path.append(os.path.join(os.path.dirname(__file__), 'lib'))
    from litellm import completion, check_valid_key
    import litellm
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
    

    def call_completion(self, messages, max_tokens=None, endpoint=None, model_name=None, provider=None, api_key=None, get_config_func=None):
        """Make a completion call with standardized parameters"""
        if get_config_func is None:
            get_config_func = self.get_config
        try:
            # Construct model string based on provider if provided
            full_model = model_name or get_config_func("model", "")
            if provider and full_model:
                full_model = f"{provider}/{full_model}"

            kwargs = {
                "model": full_model,
                "messages": messages,
                "max_tokens": max_tokens or get_config_func("extend_selection_max_tokens", 70)
            }
            # Only include api_base if endpoint is not empty
            endpoint = endpoint or get_config_func("endpoint", None)
            if isinstance(endpoint, str) and endpoint.startswith('http'):
                kwargs["api_base"] = endpoint
            api_key = api_key or get_config_func("api_key", None)
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
        add("btn_ok", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN, 
                BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})

        add("label_provider", "FixedText", HORI_MARGIN, VERT_MARGIN, label_width, LABEL_HEIGHT,
            {"Label": "Provider:", "NoLabel": True})
        add("combo_provider", "ComboBox", HORI_MARGIN, LABEL_HEIGHT + VERT_MARGIN,
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Dropdown": True})
        
        add("label_api_key", "FixedText", HORI_MARGIN, LABEL_HEIGHT + VERT_MARGIN + VERT_SEP + EDIT_HEIGHT, label_width, LABEL_HEIGHT, 
            {"Label": "API Key (Optional, use with caution):", "NoLabel": True})
        add("edit_api_key", "Edit", HORI_MARGIN, LABEL_HEIGHT*2 + VERT_MARGIN + VERT_SEP + EDIT_HEIGHT, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("api_key", "")), "EchoChar": ord('*')})
        
        add("label_model", "FixedText", HORI_MARGIN, LABEL_HEIGHT*2 + VERT_MARGIN + VERT_SEP*2 + EDIT_HEIGHT*2, label_width, LABEL_HEIGHT, 
            {"Label": "Model (Required by Ollama):", "NoLabel": True})
        add("edit_model", "Edit", HORI_MARGIN, LABEL_HEIGHT*3 + VERT_MARGIN + VERT_SEP*2 + EDIT_HEIGHT*2, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("model", ""))})
        
        add("label_endpoint", "FixedText", HORI_MARGIN, LABEL_HEIGHT*3 + VERT_MARGIN + VERT_SEP*3 + EDIT_HEIGHT*3, label_width, LABEL_HEIGHT, 
            {"Label": "Endpoint URL/Port (Local or Proxy, e.g., https://api.x.ai/v1/chat/completions for Grok):", "NoLabel": True})
        add("edit_endpoint", "Edit", HORI_MARGIN, LABEL_HEIGHT*4 + VERT_MARGIN + VERT_SEP*3 + EDIT_HEIGHT*3, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(self.get_config("endpoint", "http://127.0.0.1:5000"))})
        
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
        
        # Add provider change listener
        class ProviderChangeListener(unohelper.Base, XItemListener):
            def __init__(self, api_key_ctrl, api_key_label, endpoint_ctrl, endpoint_label):
                self.api_key_ctrl = api_key_ctrl
                self.api_key_label = api_key_label
                self.endpoint_ctrl = endpoint_ctrl
                self.endpoint_label = endpoint_label
                self.api_key_ctrl.Model.HelpText = "Leave blank if not required for your provider"
                self.endpoint_ctrl.Model.HelpText = "Leave blank to use default endpoint"

            def itemStateChanged(self, event):
                provider = event.Source.Model.Text
                if not provider:
                    return
                
                # Default to requiring key until we know otherwise
                self.api_key_ctrl.setEditable(True)
                self.api_key_ctrl.setEnable(True)
                self.api_key_ctrl.Model.HelpText = "API key (required for most providers)"
                self.api_key_ctrl.Model.BackgroundColor = 0xFFFFFF
                
                try:
                    print(f"Checking provider '{provider}' requirements...")
                    
                    # Default to requiring both key and endpoint
                    needs_key = True
                    needs_endpoint = True
                    
                    # Check for known local providers
                    local_providers = ["ollama", "openai-compatible", "vllm"]
                    if provider.lower() in local_providers:
                        needs_endpoint = True
                        self.endpoint_ctrl.Model.HelpText = "Required for local provider (e.g. http://localhost:11434)"
                        self.endpoint_ctrl.Model.BackgroundColor = 0xFFFFFF
                    else:
                        needs_endpoint = False
                        self.endpoint_ctrl.Model.HelpText = "Leave blank to use provider's default endpoint"
                        self.endpoint_ctrl.Model.BackgroundColor = 0xEEEEEE
                    
                    # Check API key requirement
                    try:
                        provider_config = litellm.utils.ProviderConfigManager.get_provider_model_info(
                            model=None,
                            provider=litellm.utils.LlmProviders(provider))
                        needs_key = provider_config.get_api_key('needed') == 'needed'
                    except:
                        needs_key = True  # Fallback to requiring key if check fails

                    if not needs_key:
                        print(f"Provider '{provider}' does NOT require API key")
                        self.api_key_ctrl.setEditable(False)
                        self.api_key_ctrl.setEnable(False)
                        self.api_key_ctrl.Model.HelpText = "This provider doesn't require an API key"
                        self.api_key_ctrl.Model.BackgroundColor = 0xEEEEEE
                    else:
                        print(f"Provider '{provider}' requires API key")
                        self.api_key_ctrl.setEditable(True)
                        self.api_key_ctrl.setEnable(True)
                        self.api_key_ctrl.Model.HelpText = "API key required"
                        self.api_key_ctrl.Model.BackgroundColor = 0xFFFFFF
                        
                    # Set endpoint field state
                    self.endpoint_ctrl.setEditable(needs_endpoint)
                    self.endpoint_ctrl.setEnable(needs_endpoint)
                    
                except Exception as e:
                    print(f"Error checking key requirement for {provider}: {str(e)}")

            def disposing(self, event):
                pass

        # Get all dialog controls first
        combo_provider = dialog.getControl("combo_provider")
        api_key_label = dialog.getControl("label_api_key")
        edit_api_key = dialog.getControl("edit_api_key")
        edit_endpoint = dialog.getControl("edit_endpoint")
        endpoint_label = dialog.getControl("label_endpoint")
        
        # Get known providers from LiteLLM
        try:
            providers = sorted(litellm.provider_list)
        except:
            providers = ["openai", "ollama", "anthropic", "cohere", "huggingface", "replicate"]
        
        current_provider = str(self.get_config("provider", ""))
        combo_provider.addItems(providers, 0)
        if current_provider in providers:
            combo_provider.Model.Text = current_provider

        # Set up provider change listener
        provider_listener = ProviderChangeListener(
            edit_api_key, 
            api_key_label,
            edit_endpoint,
            endpoint_label
        )
        combo_provider.addItemListener(provider_listener)
        
        # Initialize state based on current provider
        if current_provider:
            # Create proper ItemEvent structure
            item_event = uno.createUnoStruct("com.sun.star.awt.ItemEvent")
            item_event.Source = combo_provider
            item_event.Selected = 0  # Default selection index
            provider_listener.itemStateChanged(item_event)
        
        edit_api_key = dialog.getControl("edit_api_key")
        edit_api_key.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("api_key", "")))))
        
        edit_model = dialog.getControl("edit_model")
        edit_model.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("model", "")))))
        
        edit_endpoint = dialog.getControl("edit_endpoint")
        edit_endpoint.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("endpoint", "http://127.0.0.1:5000")))))
        
        edit_extend_selection_max_tokens = dialog.getControl("edit_extend_selection_max_tokens")
        edit_extend_selection_max_tokens.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("extend_selection_max_tokens", "70")))))
        
        edit_extend_selection_system_prompt = dialog.getControl("edit_extend_selection_system_prompt")
        edit_extend_selection_system_prompt.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("extend_selection_system_prompt", "")))))
        
        edit_edit_selection_max_new_tokens = dialog.getControl("edit_edit_selection_max_new_tokens")
        edit_edit_selection_max_new_tokens.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("edit_selection_max_new_tokens", "0")))))
        
        edit_edit_selection_system_prompt = dialog.getControl("edit_edit_selection_system_prompt")
        edit_edit_selection_system_prompt.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("edit_selection_system_prompt", "")))))
        
        combo_provider.setFocus()

        # Add listener for test button using a UNO-compatible ActionListener
        from com.sun.star.awt import XActionListener
        class TestConnectionListener(unohelper.Base, XActionListener):
            def __init__(self, endpoint_ctrl, model_ctrl, provider_ctrl, api_key_ctrl, result_ctrl, main_job, api_key_label):
                self.endpoint_ctrl = endpoint_ctrl
                self.model_ctrl = model_ctrl
                self.provider_ctrl = provider_ctrl
                self.api_key_ctrl = api_key_ctrl
                self.result_ctrl = result_ctrl
                self.main_job = main_job
                self.api_key_label = api_key_label

            def disposing(self, source):
                pass

            def actionPerformed(self, event):
                endpoint = self.endpoint_ctrl.getModel().Text
                model_name = self.model_ctrl.getModel().Text
                provider = self.provider_ctrl.getModel().Text
                api_key = self.api_key_ctrl.getModel().Text
                try:
                    # Turn on debug for this test call
                    litellm._turn_on_debug()

                    response = self.main_job.call_completion(
                        messages=[{"role": "user", "content": "Hello, are you working?"}],
                        max_tokens=10,
                        endpoint=endpoint,
                        model_name=model_name,
                        provider=provider,
                        api_key=api_key,
                        get_config_func=lambda x,y: '' # disable pulling anything from config since we're testing config values now
                    )
                    self.result_ctrl.setText("Success: " + response.choices[0].message.content)
                except Exception as e:
                    self.result_ctrl.setText("Failed: " + str(e))
                    print(f"Test Connection Error: {str(e)}")

        btn_test = dialog.getControl("btn_test")
        test_listener = TestConnectionListener(
            edit_endpoint,
            edit_model, 
            combo_provider,
            edit_api_key,
            dialog.getControl("test_result"),
            self,
            api_key_label
        )
        btn_test.addActionListener(test_listener)

        if dialog.execute():
            result = {
                "endpoint": edit_endpoint.getModel().Text,
                "model": edit_model.getModel().Text,
                "provider": combo_provider.Model.Text,
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
