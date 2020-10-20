from configparser import ConfigParser

# Initialize the Parser.
config = ConfigParser()

# Add the Section.
config.add_section('graph_api')

# Set the Values.
config.set('graph_api', 'client_id', '')
config.set('graph_api', 'client_secret', '')
config.set('graph_api', 'redirect_uri', '')

# Write the file.
with open(file='samples/configs/config_video.ini', mode='w+') as f:
    config.write(f)
