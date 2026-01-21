"""
Azure AI Configuration Loader

Shared configuration for all tools that use Azure AI services.

Copyright 2025-2026 Andre Lorbach
Licensed under Apache License 2.0
"""

import os
from pathlib import Path

# YAML support (optional)
try:
    import yaml
    YAML_AVAILABLE = True
except ImportError:
    YAML_AVAILABLE = False


class AzureAIConfig:
    """Shared Azure AI configuration loader."""
    
    CONFIG_FILE = "config/azure_ai.yaml"
    
    def __init__(self, root_dir=None):
        """
        Initialize Azure AI configuration.
        
        Args:
            root_dir: Root directory of the project. If None, auto-detected.
        """
        if root_dir:
            self.root_dir = Path(root_dir)
        else:
            # Auto-detect: go up from this file's location
            self.root_dir = Path(__file__).parent.parent.parent
        
        self.config = self._load_config()
    
    def _load_config(self):
        """Load configuration from file and environment variables."""
        # Default configuration
        config = {
            'azure_openai': {
                'endpoint': '',
                'api_key': '',
                'api_version': '2024-02-15-preview',
                'deployment_name': 'gpt-4'
            },
            'azure_document_intelligence': {
                'endpoint': '',
                'api_key': ''
            },
            'settings': {
                'prefer_env_vars': True,
                'timeout': 60,
                'max_retries': 3
            }
        }
        
        # Try to load from config file
        config_path = self.root_dir / self.CONFIG_FILE
        if config_path.exists() and YAML_AVAILABLE:
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    file_config = yaml.safe_load(f)
                    if file_config:
                        self._merge_config(config, file_config)
                print(f"[INFO] Loaded Azure AI config from {config_path}")
            except Exception as e:
                print(f"[WARNING] Could not load Azure AI config file: {e}")
        elif config_path.exists() and not YAML_AVAILABLE:
            print("[WARNING] PyYAML not installed. Cannot load config file. Install with: pip install pyyaml")
        
        # Override with environment variables if preferred
        if config['settings'].get('prefer_env_vars', True):
            self._load_env_vars(config)
        
        return config
    
    def _merge_config(self, base, override):
        """Recursively merge override into base config."""
        for key, value in override.items():
            if key in base and isinstance(base[key], dict) and isinstance(value, dict):
                self._merge_config(base[key], value)
            else:
                base[key] = value
    
    def _load_env_vars(self, config):
        """Load API keys from environment variables."""
        # Azure OpenAI
        if os.environ.get('AZURE_OPENAI_ENDPOINT'):
            config['azure_openai']['endpoint'] = os.environ['AZURE_OPENAI_ENDPOINT']
        if os.environ.get('AZURE_OPENAI_API_KEY'):
            config['azure_openai']['api_key'] = os.environ['AZURE_OPENAI_API_KEY']
        if os.environ.get('AZURE_OPENAI_DEPLOYMENT'):
            config['azure_openai']['deployment_name'] = os.environ['AZURE_OPENAI_DEPLOYMENT']
        if os.environ.get('AZURE_OPENAI_API_VERSION'):
            config['azure_openai']['api_version'] = os.environ['AZURE_OPENAI_API_VERSION']
        
        # Azure Document Intelligence
        if os.environ.get('AZURE_DOC_INTEL_ENDPOINT'):
            config['azure_document_intelligence']['endpoint'] = os.environ['AZURE_DOC_INTEL_ENDPOINT']
        if os.environ.get('AZURE_DOC_INTEL_API_KEY'):
            config['azure_document_intelligence']['api_key'] = os.environ['AZURE_DOC_INTEL_API_KEY']
    
    def save_config(self):
        """Save current configuration to file."""
        if not YAML_AVAILABLE:
            print("[ERROR] PyYAML not installed. Cannot save config file.")
            return False
        
        config_path = self.root_dir / self.CONFIG_FILE
        
        # Create config directory if it doesn't exist
        config_path.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            # Don't save API keys if they came from environment variables
            save_config = {
                'azure_openai': {
                    'endpoint': self.config['azure_openai']['endpoint'],
                    'api_key': '',  # Don't save API keys
                    'api_version': self.config['azure_openai']['api_version'],
                    'deployment_name': self.config['azure_openai']['deployment_name']
                },
                'azure_document_intelligence': {
                    'endpoint': self.config['azure_document_intelligence']['endpoint'],
                    'api_key': ''  # Don't save API keys
                },
                'settings': self.config['settings']
            }
            
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.safe_dump(save_config, f, default_flow_style=False, indent=2)
            
            print(f"[INFO] Saved Azure AI config to {config_path}")
            return True
            
        except Exception as e:
            print(f"[ERROR] Failed to save Azure AI config: {e}")
            return False
    
    # Properties for easy access
    @property
    def openai_endpoint(self):
        return self.config['azure_openai']['endpoint']
    
    @openai_endpoint.setter
    def openai_endpoint(self, value):
        self.config['azure_openai']['endpoint'] = value
    
    @property
    def openai_api_key(self):
        return self.config['azure_openai']['api_key']
    
    @openai_api_key.setter
    def openai_api_key(self, value):
        self.config['azure_openai']['api_key'] = value
    
    @property
    def openai_deployment(self):
        return self.config['azure_openai']['deployment_name']
    
    @openai_deployment.setter
    def openai_deployment(self, value):
        self.config['azure_openai']['deployment_name'] = value
    
    @property
    def openai_api_version(self):
        return self.config['azure_openai']['api_version']
    
    @openai_api_version.setter
    def openai_api_version(self, value):
        self.config['azure_openai']['api_version'] = value
    
    @property
    def doc_intel_endpoint(self):
        return self.config['azure_document_intelligence']['endpoint']
    
    @doc_intel_endpoint.setter
    def doc_intel_endpoint(self, value):
        self.config['azure_document_intelligence']['endpoint'] = value
    
    @property
    def doc_intel_api_key(self):
        return self.config['azure_document_intelligence']['api_key']
    
    @doc_intel_api_key.setter
    def doc_intel_api_key(self, value):
        self.config['azure_document_intelligence']['api_key'] = value
    
    @property
    def timeout(self):
        return self.config['settings'].get('timeout', 60)
    
    @property
    def max_retries(self):
        return self.config['settings'].get('max_retries', 3)
    
    def is_openai_configured(self):
        """Check if Azure OpenAI is properly configured."""
        return bool(self.openai_endpoint and self.openai_api_key)
    
    def is_doc_intel_configured(self):
        """Check if Azure Document Intelligence is properly configured."""
        return bool(self.doc_intel_endpoint and self.doc_intel_api_key)
    
    def get_status_text(self):
        """Get a human-readable status of the configuration."""
        lines = []
        
        if self.is_openai_configured():
            lines.append(f"✓ Azure OpenAI: Configured ({self.openai_endpoint})")
        else:
            missing = []
            if not self.openai_endpoint:
                missing.append("endpoint")
            if not self.openai_api_key:
                missing.append("API key")
            lines.append(f"✗ Azure OpenAI: Not configured (missing: {', '.join(missing)})")
        
        if self.is_doc_intel_configured():
            lines.append(f"✓ Document Intelligence: Configured ({self.doc_intel_endpoint})")
        else:
            missing = []
            if not self.doc_intel_endpoint:
                missing.append("endpoint")
            if not self.doc_intel_api_key:
                missing.append("API key")
            lines.append(f"✗ Document Intelligence: Not configured (missing: {', '.join(missing)})")
        
        return '\n'.join(lines)


# Singleton instance for easy import
_config_instance = None

def get_azure_config(root_dir=None):
    """Get the shared Azure AI configuration instance."""
    global _config_instance
    if _config_instance is None:
        _config_instance = AzureAIConfig(root_dir)
    return _config_instance
