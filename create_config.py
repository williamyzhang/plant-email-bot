import yaml
import io

data = {
    'from_email': 'your_email_here@gmail.com',
    'to_email': 'email_to_send_to@example.com',
    'ignore_means_yes': False,
    'spreadsheet_name': 'Plant Task Data',
    'initiated': False
}

with io.open('config.yaml', 'w', encoding='utf8') as outfile:
    yaml.dump(data, outfile, default_flow_style=False, allow_unicode=True,
              sort_keys=False)