"""
Main entry point for the RTR (Registratie Toepasbare Regels) archiving application.

This script processes command-line arguments and initiates the archiving process
for RTR data from the IPLO Omgevingswet API.
"""

from rtr import RTR
from commands import ArgumentParser  

def main():
    """Execute the RTR archiving process with parsed command-line arguments."""
    args = ArgumentParser.parse_command_line_arguments()
    RTR(args).archive_activities()

if __name__ == "__main__":
    main()
