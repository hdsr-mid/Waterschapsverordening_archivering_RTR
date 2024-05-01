import os
from datetime import datetime
import requests
import argparse
import urllib.parse

from excel import ExcelHandler
from vendor import Vendor

class RTR:
    def __init__(self, software):
        self.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  
        self.args = self.parse_command_line_arguments()
        self.api_key = self.load_api_key(os.path.join(self.base_dir, 'code', f"{self.args.env}_API_key.txt"))
        self.headers = {'Accept': 'application/hal+json, application/xml', 'x-api-key': self.api_key}
        self.base_url = self.compose_base_url(self.args.env)
        self.vendor = Vendor(software, self.args.env)
        self.urns = self.vendor.urns
        self.geo_variables = self.vendor.variable_names_by_index
        self.session = requests.Session()
        self.sttr_url_per_activity = {}
        self.werkingsgebied_per_activity = {}
        self.excel_handler = ExcelHandler(self.base_dir, self.args.env, self.args.date)
        self.run_once = False
        
    @staticmethod
    def parse_command_line_arguments():
        parser = argparse.ArgumentParser(description="Process some environment settings and actions.")
        parser.add_argument('--env', type=str, default="prod", choices=['prod', 'pre'],
                            help='Environment setting: prod (default) or pre.')
        parser.add_argument('--date', type=str, default=datetime.now().strftime("%d-%m-%Y"),
                            help='Date in the format dd-mm-yyyy, default is today\'s date.')
        parser.add_argument('--sttr', action='store_true',
                            help='Flag to log sttr files in .xml if present.')
        args = parser.parse_args()
        return args

    @staticmethod
    def load_api_key(api_key_file):
        with open(api_key_file) as key_file:
            return key_file.read().strip()

    def archive_activities(self):
        for row, activity in enumerate(self.urns, 2):
            self.process_activity(activity, row)
        if self.args.sttr: 
            self.archive_sttr_files()
        self.excel_handler.close_workbook()

    def process_activity(self, activity, row):
        name, _, uri, _, activity_group, rule_reference, _ = activity
        response_json = self.get_activity_data(uri)
        if response_json:
            self.archive_activity_data(
                row, name, uri, activity_group, rule_reference, response_json
            )

    def get_activity_data(self, uri):
        url = self.compose_activity_url(uri)
        response = self.session.get(url, headers=self.headers)
        
        if response.ok:
            self.update_werkingsgebied_per_activity(response.json())
            return response.json()
        print(f"Error fetching data for URI {uri}: {response.status_code}")
        return None
    

    def update_werkingsgebied_per_activity(self, json_data):
        activity_description = json_data.get('omschrijving', 'No description')
        identifications = [loc['identificatie'] for loc in json_data.get('locaties', [])]

        # Initialize a list to hold the descriptions matched from self.geo_variables or the specific case
        matched_descriptions = []
        for url in identifications:
            if url == 'nl.imow-ws0636.ambtsgebied.HDSR':
                description = 'Ambtsgebied HDSR'
            else:
                # Extract the last two digits of the identifier to get the index
                index = url.split('.')[-1][-2:]  # Assumes format ends with two digits like '2023000038'
                description = self.geo_variables.get(index, f"null: {url}")  # Get the description or 'Unknown description'
            matched_descriptions.append(description)
        
        # Check if the activity description already exists in the dictionary
        if activity_description in self.werkingsgebied_per_activity:
            self.werkingsgebied_per_activity[activity_description].extend(matched_descriptions)
        else:
            self.werkingsgebied_per_activity[activity_description] = matched_descriptions

        file_path = os.path.join(self.base_dir, 'log', "werkingsgebieden.txt")
        with open(file_path, 'w') as file:
            for key, values in self.werkingsgebied_per_activity.items():
                file.write(f"{key}: {', '.join(values)}\n\n")




    @staticmethod
    def extract_werkzaamheden(data):
        werkzaamheden_list = []
        if "werkzaamheden" in data["_links"]:
            for werkzaamheid in data["_links"]["werkzaamheden"]:
                extracted_id = werkzaamheid["href"].split("/")[(-1)]
                werkzaamheden_list.append(extracted_id)
        return [', '.join(werkzaamheden_list)] if werkzaamheden_list else [""]

    def fetch_and_process_changes(self, data):
        urn_name = data["urn"].split(".")[-1]
        changes = ["", "", "", ""]
        
        if "regelBeheerObjecten" in data:
            for object in data["regelBeheerObjecten"]:
                object_type, last_changed = self.process_individual_object(urn_name, object)
                if object_type in {"Conclusie", "Melding", "Aanvraag vergunning", "Informatie"}:
                    index = ["Conclusie", "Melding", "Aanvraag vergunning", "Informatie"].index(object_type)
                    changes[index] = last_changed
        return changes

    def process_individual_object(self, urn_name, object):
        object_type = object["typering"]
        if object_type == "Indieningsvereisten":
            object_type = object["toestemming"]["waarde"]

        functional_structure_reference = object["functioneleStructuurRef"]
        last_changed = self.get_regelbeheerobject(urn_name, object_type, functional_structure_reference)
        return object_type, last_changed
    
    def get_regelbeheerobject(self, urn_name, object_type, functional_structure_reference):
        url = self.compose_regel_beheer_object_url(functional_structure_reference)
        response = self.session.get(url, headers=self.headers)

        if response.ok:
            data = response.json()
            self.append_sttr_file(urn_name, object_type, data)
            last_changed = self.get_last_change_date(data)
            return last_changed

    def get_last_change_date(self, data):
        embedded = data.get('_embedded', {})
        applicable_rules = embedded.get('toepasbareRegels', [])
        if applicable_rules:
            return applicable_rules[0].get("laatsteWijzigingDatum", "")
        else:
            return ""
    
    def append_sttr_file(self, urn_name, regelbeheerobject_type, data):
        try:
            sttr_bestand_href = data['_embedded']['toepasbareRegels'][0]['_links']['sttrBestand']['href']
            if regelbeheerobject_type != "null":
                regelbeheerobject_name = urn_name + "_" + regelbeheerobject_type.replace(" ", "_")
                self.sttr_url_per_activity[regelbeheerobject_name] = sttr_bestand_href
            
        except KeyError as e:
            identifier = self.extract_identifier(data)
            print(f"Data missing key: '{e}'. Regelbeheerobject: {identifier}")

    def extract_identifier(self, data):
        try:
            url = data.get('_links', {}).get('self', {}).get('href', "")
            functionele_structuur_ref = urllib.parse.parse_qs(urllib.parse.urlparse(url).query).get('functioneleStructuurRef', [''])[0]
            return functionele_structuur_ref.split('/')[-1]
        except Exception:
            return "Unknown"  

    @staticmethod
    def compose_base_url(env):
        if env == "prod":
            return "https://service.omgevingswet.overheid.nl/publiek/toepasbare-regels/api"
        if env == "pre":
            return "https://service.pre.omgevingswet.overheid.nl/publiek/toepasbare-regels/api"
        raise ValueError("Invalid environment specified")

    def compose_activity_url(self, uri):
        return f"{self.base_url}/rtrgegevens/v2/activiteiten/{uri}?datum={self.args.date}"

    def compose_regel_beheer_object_url(self, functional_structure_reference):
        return f"{self.base_url}/toepasbareregelsuitvoerengegevens/v1/toepasbareRegels?functioneleStructuurRef={functional_structure_reference}&datum={self.args.date}"

    def archive_activity_data(self, row, name, uri, activity_group, rule_reference, data):
        werkzaamheden = self.extract_werkzaamheden(data)
        
        changes = self.fetch_and_process_changes(data)
        data_to_write = [name, uri, activity_group, rule_reference] + werkzaamheden + changes
        self.excel_handler.write_data_to_cells(row, data_to_write)

    def archive_sttr_files(self):
        for key, url in self.sttr_url_per_activity.items():
            identifier = url.split('/toepasbareRegels/')[1].split('/')[0]
            response = self.session.get(url, headers=self.headers)
                         
            if response.status_code == 200:
                with open(os.path.join(self.base_dir, 'log', f'STTR_RegelBeheerObjecten/STTR_{identifier}_{key}.xml'), 'w', encoding='utf-8') as file:
                    file.write(response.text)
            else:
                print(f"Failed to download data from {url}, status code: {response.status_code}")










    # def fetch_location_details(self):
    #     # Adjust the URL based on the specific object you are querying

    #     #activiteitlocatieaanduidingen/_zoek
    #     #gebiedsaanwijzingen/_zoek
    #     #regelteksten/_zoek
    #     #omgevingsnormen/_zoek
    #     #omgevingswaarden/_zoek

    #     gebieds_url = "https://service.omgevingswet.overheid.nl/publiek/omgevingsdocumenten/api/presenteren/v7/gebiedsaanwijzingen/_zoek"
    #     headers = self.headers
    #     headers['Content-Type'] = 'application/json'  # Ensure header includes Content-Type as application/json

    #     # The body of the POST request with the specified zoekParameters
    #     search_payload = {
    #         "zoekParameters": [
    #             {
    #                 "parameter": "locatie.identificatie",
    #                 "zoekWaarden": [
    #                     "nl.imow-ws0636.gebied.2023000034"
    #                     ]
    #             }
    #         ]
    #     }

    #     # Making the POST request
    #     response = self.session.post(gebieds_url, headers=headers, json=search_payload)
    #     if response.ok:
    #         print('ok', response.json(), '\n')
    #     else:
    #         print(f"Failed to fetch data: {response.status_code} {response.text}")
