"""
RTR (Registratie Toepasbare Regels) data archiving module.

This module handles the main archiving logic for RTR data from the IPLO Omgevingswet API.
It processes activities, extracts werkzaamheden and regelbeheerobjecten, and generates
Excel reports and optional STTR XML files.
"""

import os
import requests
import urllib.parse
from collections import OrderedDict
from datetime import datetime

from excel import ExcelHandler
from powerbi import PowerBIData
from commands import ArgumentParser  

class RTR:
    """
    Main RTR archiving class.
    
    This class handles the complete archiving workflow:
    - Fetching activity data from the IPLO RTR API
    - Processing werkzaamheden and regelbeheerobjecten
    - Generating Excel reports with color-coded dates
    - Optionally archiving STTR DMN files
    - Optionally exporting werkingsgebieden mappings
    """
    
    def __init__(self, args):
        self.args = args 
        self.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) 
        self.api_key = self.load_api_key(os.path.join(self.base_dir, 'data', f"{self.args.env}_API_key.txt"))
        self.base_url = self.compose_base_url(self.args.env)
        self.powerbi_env = "PRE" if self.args.env == "pre" else "PROD"
        self.session_headers = {'Accept': 'application/hal+json, application/xml', 'x-api-key': self.api_key}
        self.urn_file_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                                 'data', 
                                 f"A1. Welke activiteiten zijn gewijzigd {self.powerbi_env}.xlsx")
        self.location_file_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                                'data', 
                                f"A3. Wie gebruikt welke locaties (in STTR) {self.powerbi_env}.xlsx")
        self.headers = ["Activiteit                   ",
                        "Urn",              
                        "Aantal",
                        "Werkzaamheden",          
                        "Wijziging Conclusie",
                        "Wijziging Melding",
                        "Wijziging Aanvraag vergunning",
                        "Wijziging Informatie"  ]
        self.powerbi = PowerBIData(self.urn_file_path, self.location_file_path)
        self.urns = self.powerbi.get_urns(self.args.overheid) 
        self.geo_variables = self.powerbi.get_location_identifiers(self.args.overheid)
        self.session = requests.Session()
        self.sttr_url_per_activity = {}
        self.werkingsgebied_per_activity = {}
        self.unique_werkingsgebieden = set()
        self.excel_handler = None

    @staticmethod
    def load_api_key(api_key_file):
        with open(api_key_file) as key_file:
            return key_file.read().strip()

    def archive_activities(self):
        for activity in self.urns:
            print(activity)
            self.collect_unique_werkingsgebieden(activity)
        
        headers = self.headers + sorted(self.unique_werkingsgebieden)
        self.excel_handler = ExcelHandler(self.args.overheid, self.base_dir, self.args.env, self.args.date, headers)
        
        for row, activity in enumerate(self.urns, 2):
            self.process_activity(activity, row)
        
        self.excel_handler.close_workbook()
        if self.args.sttr:
            self.archive_sttr_files()
        if self.args.location:
            self.sorted_gebied_to_activities = self.invert_werkingsgebied_mapping()
            self.write_werkingsgebieden_to_file()

    def collect_unique_werkingsgebieden(self, activity):
        government, name, uri = activity
        response_json = self.get_activity_data(uri)
        if response_json:
            self.update_werkingsgebied_per_activity(response_json)

    def process_activity(self, activity, row):
        government, name, uri = activity
        response_json = self.get_activity_data(uri)
        
        if response_json:
            self.archive_activity_data(row, name, uri, response_json)

    def get_activity_data(self, uri):
        url = self.compose_activity_url(uri)
        response = self.session.get(url, headers=self.session_headers)
        
        if response.ok:
            self.update_werkingsgebied_per_activity(response.json())
            return response.json()
        print(f"Error fetching data for URI {uri}: {response.status_code}")
        return None
    
    def update_werkingsgebied_per_activity(self, json_data):
        activity_description = self.extract_activity_description(json_data)
        identifications = self.extract_identifications(json_data)
        matched_descriptions = self.match_descriptions(identifications)
        self.update_activity_mapping(activity_description, matched_descriptions)

    def extract_activity_description(self, json_data):
        return json_data.get('omschrijving', 'No description')

    def extract_identifications(self, json_data):
        return [loc['identificatie'] for loc in json_data.get('locaties', [])]

    def match_descriptions(self, identifications):
        matched_descriptions = []
        for url in identifications:
            description = self.get_description(url)
            matched_descriptions.append(description)
        return matched_descriptions

    def get_description(self, url):
        if self.geo_variables.get(url) != None:
            return self.geo_variables.get(url)
        elif "ambtsgebied" in url:
            return 'Ambtsgebied'
        else:
            return url
            
    def update_activity_mapping(self, activity_description, matched_descriptions):
        self.unique_werkingsgebieden.update(matched_descriptions)
        if activity_description in self.werkingsgebied_per_activity:
            self.werkingsgebied_per_activity[activity_description].extend(matched_descriptions)
        else:
            self.werkingsgebied_per_activity[activity_description] = matched_descriptions
        
    def invert_werkingsgebied_mapping(self):
        gebied_to_activities = {}
        for activity, gebieden in self.werkingsgebied_per_activity.items():
            for gebied in gebieden:
                if gebied not in gebied_to_activities:
                    gebied_to_activities[gebied] = set()  # Use a set to avoid duplicates
                gebied_to_activities[gebied].add(activity)
        
        # Convert the sets back to lists and sort them
        for gebied in gebied_to_activities:
            gebied_to_activities[gebied] = sorted(gebied_to_activities[gebied])
        
        sorted_gebied_to_activities = OrderedDict(sorted(gebied_to_activities.items()))
        return sorted_gebied_to_activities

    def write_werkingsgebieden_to_file(self):      
        file_path = self.create_file_path('log', f"{self.args.overheid.replace(' ', '_')}_waterschapsverordening_werkingsgebieden_{self.args.env}_{self.args.date}.txt")
        gebied_to_activities = self.invert_werkingsgebied_mapping()
        self.write_to_file(file_path, gebied_to_activities)

    def create_file_path(self, directory, filename):
        # Ensure the directory exists
        dir_path = os.path.join(self.base_dir, directory)
        os.makedirs(dir_path, exist_ok=True)
        
        # Create the full file path
        file_path = os.path.join(dir_path, filename)
        return file_path

    def write_to_file(self, file_path, data):
        with open(file_path, 'w', encoding='utf-8') as file:
            for key, values in data.items():
                file.write(f"{key}:\n")
                for value in values:
                    file.write(f"  - {value}\n")
                file.write("\n")

    @staticmethod
    def extract_werkzaamheden(data):
        werkzaamheden_list = []
        counter = 0
        if "werkzaamheden" in data["_links"]:
            for werkzaamheid in data["_links"]["werkzaamheden"]:
                extracted_id = werkzaamheid["href"].split("/")[(-1)]
                clean_id = extracted_id.split("?")[0]
                werkzaamheden_list.append(clean_id)
                counter += 1
        return [counter, [', '.join(werkzaamheden_list)] if werkzaamheden_list else [""]]

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
        response = self.session.get(url, headers=self.session_headers)

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

    def archive_sttr_files(self):
        folder_name = self.excel_handler.generate_file_path(self.args.date, self.args.overheid, self.args.env, "DMN-files", self.base_dir)
        os.makedirs(folder_name, exist_ok=True)
            
        for key, url in self.sttr_url_per_activity.items():
            identifier = url.split('/toepasbareRegels/')[1].split('/')[0]
            response = self.session.get(url, headers=self.session_headers)
                
            if response.status_code == 200:
                file_path = os.path.join(folder_name, f'STTR_{identifier}_{key}.xml')
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(response.text)
            else:
                print(f"Failed to download data from {url}, status code: {response.status_code}")

    def extract_identifier(self, data):
        try:
            url = data.get('_links', {}).get('self', {}).get('href', "")
            functionele_structuur_ref = urllib.parse.parse_qs(urllib.parse.urlparse(url).query).get('functioneleStructuurRef', [''])[0]
            return functionele_structuur_ref.split('/')[-1]
        except Exception:
            return "Unknown"  

    @staticmethod
    def compose_base_url(env):
        return f"https://service{'' if env == 'prod' else '.pre'}.omgevingswet.overheid.nl/publiek/toepasbare-regels/api"

    def compose_activity_url(self, uri):
        return f"{self.base_url}/rtrgegevens/v2/activiteiten/{uri}?datum={self.args.date}"

    def compose_regel_beheer_object_url(self, functional_structure_reference):
        return f"{self.base_url}/toepasbareregelsuitvoerengegevens/v1/toepasbareRegels?functioneleStructuurRef={functional_structure_reference}&datum={self.args.date}"

    def archive_activity_data(self, row, name, uri, data):
        werkzaamheden = self.extract_werkzaamheden(data)
        changes = self.fetch_and_process_changes(data)

        unique_werkingsgebieden = sorted(self.unique_werkingsgebieden)
        werkingsgebieden_indices = {gebied: index for index, gebied in enumerate(unique_werkingsgebieden)}
        activity_werkingsgebieden_presence = [" "] * len(unique_werkingsgebieden)

        for gebied in self.werkingsgebied_per_activity.get(name.lower(), []):
            activity_werkingsgebieden_presence[werkingsgebieden_indices[gebied]] = 1
            
        data_to_write = [name, uri, werkzaamheden[0]] + werkzaamheden[1] + changes + activity_werkingsgebieden_presence
        self.excel_handler.write_data_to_cells(row, data_to_write)                 
        
    def decodeSpecialChar(self, string):
        return string.encode("latin1").decode("utf-8")
