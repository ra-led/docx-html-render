import yaml

def load_config(config_path='./conf.yaml'):
    """
    Load configuration from a YAML file.

    :param config_path: Path to the YAML configuration file.
    :return: Dictionary containing the configuration settings.
    """
    with open(config_path, 'r') as file:
        config = yaml.safe_load(file)
    return config


CONF = load_config()
