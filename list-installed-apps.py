import os
import sys
import winreg
import argparse
import csv


def list_installed_software(filter_keyword=None, output_path=None, exclude_keyword=None):
    software_details = []
    registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                  r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", 0,
                                  winreg.KEY_READ)

    try:
        count_subkey = winreg.QueryInfoKey(registry_key)[0]
        for i in range(count_subkey):
            software_name = None
            publisher = None
            version = None
            description = None

            try:
                key_name = winreg.EnumKey(registry_key, i)
                key_path = os.path.join(r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", key_name)
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ) as subkey:
                    software_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                    publisher = get_registry_value(subkey, "Publisher")
                    version = get_registry_value(subkey, "DisplayVersion")
                    description = get_registry_value(subkey, "Description")

            except EnvironmentError:
                continue

            if software_name:
                include_software = True

                # If there's a filter keyword, only include if it matches the publisher
                if filter_keyword:
                    include_software = publisher and filter_keyword.lower() in publisher.lower()

                # If there's an exclude keyword, exclude if it matches the publisher
                if exclude_keyword:
                    include_software = not (publisher and exclude_keyword.lower() in publisher.lower())

                if include_software:
                    software_details.append((software_name, version, publisher, description))


    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        winreg.CloseKey(registry_key)

    if output_path:
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(['Software Name', 'Version', 'Publisher', 'Description'])  # CSV header
            csv_writer.writerows(software_details)
        print(f"Output saved to {output_path}")
    else:
        for software in software_details:
            print(software)


def get_registry_value(subkey, value_name):
    """Utility function to fetch registry value and handle exceptions."""
    try:
        return winreg.QueryValueEx(subkey, value_name)[0]
    except EnvironmentError:
        return None


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="List installed software with optional filtering by publisher.")

    # Argument for including software by publisher's name
    parser.add_argument('-f', '--filter', type=str,
                        help="Optional keyword to include software by publisher. For example, '-f Microsoft' would include software with 'Microsoft' as the publisher.")

    # Argument for excluding software by publisher's name
    parser.add_argument('-e', '--exclude', type=str,
                        help="Optional keyword to exclude software by publisher. For example, '-e Microsoft' would exclude software with 'Microsoft' as the publisher.")

    # Argument for specifying an output file
    parser.add_argument('-o', '--output', type=str,
                        help="Optional path to save the output. If not specified, the software list will be printed to the console. For example, '-o C:\\path\\to\\output.txt' would save the list to the specified file.")

    args = parser.parse_args()

    # Logic to prevent conflicts between filter and exclude
    if args.filter and args.exclude:
        print("You cannot use both filter and exclude options at the same time.")
        sys.exit(1)

    list_installed_software(args.filter, args.output, args.exclude)

