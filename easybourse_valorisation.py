import os
import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import logging
import re

# Logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class EasyBourseValorisationDownloader:
    def __init__(self, username, password, download_dir=None):
        self.username = username
        self.password = password
        self.base_url = "https://www.easybourse.com"
        self.valorisation_url = f"{self.base_url}/secure/compte/valorisation"

        # Define download directory
        if download_dir is None:
            self.download_dir = os.path.dirname(os.path.abspath(__file__))
        else:
            self.download_dir = download_dir

        logger.info(f"Download directory: {self.download_dir}")

    def setup_driver(self):
        """Configure and return a Selenium driver with automatic download"""
        try:
            options = Options()

            # Configuration for automatic download
            prefs = {
                "download.default_directory": self.download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "plugins.always_open_pdf_externally": True  # For PDFs if needed
            }
            options.add_experimental_option("prefs", prefs)

            # ===== HEADLESS MODE ENABLED =====
            options.add_argument('--headless=new')  # New more stable headless mode
            options.add_argument('--disable-gpu')  # Recommended for headless mode
            options.add_argument('--window-size=1920,1080')  # Important in headless

            # Options to avoid detection
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)

            # Other useful options
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-web-security')
            options.add_argument('--disable-features=VizDisplayCompositor')
            options.add_argument('--disable-extensions')

            # To reduce logs
            options.add_argument('--log-level=3')  # Fatal errors only
            options.add_argument('--silent')

            driver = webdriver.Chrome(options=options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

            # No need for set_window_size as already defined in options

            return driver
        except Exception as e:
            logger.error(f"Error configuring driver: {e}")
            raise

    def login(self, driver):
        """Log in to EasyBourse"""
        try:
            # Step 1: Login page - Enter username
            logger.info("Navigating to login page...")
            driver.get(f"{self.base_url}/login")

            # Wait for username field to be visible
            wait = WebDriverWait(driver, 10)
            username_field = wait.until(
                EC.presence_of_element_located((By.NAME, "username"))
            )

            # Accept cookies if present
            try:
                logger.info("Accepting cookies...")
                time.sleep(2)
                driver.find_element(By.XPATH, "//button[contains(text(), 'Ok pour moi')]").click()
            except:
                pass

            logger.info("Entering username...")
            username_field.send_keys(self.username)

            # Click Continue button
            continue_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Continuer')]")
            continue_button.click()

            # Step 2: Password page
            logger.info("Waiting for password page...")
            time.sleep(2)

            # Handle virtual keyboard or normal field
            try:
                # Detect virtual keyboard
                virtual_keyboard = None
                for key_id in range(1000):
                    try:
                        elements = driver.find_elements(By.CLASS_NAME, f"jss{key_id}")
                        digits = [el.text.strip() for el in elements if
                                  el.text.strip().isdigit() and len(el.text.strip()) == 1]
                        if set(digits) == set("0123456789"):
                            virtual_keyboard = elements
                            logger.info(f"✅ Virtual keyboard detected with jss{key_id}")
                            break
                    except:
                        continue

                if virtual_keyboard:
                    logger.info("Using virtual keyboard...")
                    for digit in self.password:
                        for button in virtual_keyboard:
                            if button.text == digit:
                                button.click()
                                time.sleep(0.2)
                                break
                else:
                    # Try normal password field
                    try:
                        password_field = driver.find_element(By.NAME, "password")
                        password_field.send_keys(self.password)
                    except:
                        logger.info("Please enter password manually...")
                        time.sleep(30)

            except Exception as e:
                logger.error(f"Error entering password: {e}")

            # Click Login button
            login_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Se connecter')]")
            login_button.click()

            logger.info("Logging in...")
            time.sleep(5)

            return True

        except Exception as e:
            logger.error(f"Login error: {e}")
            return False

    def download_valorisation_csv(self, driver):
        """Download the valuation CSV file"""

        # URL for CSV download page
        driver.get('https://www.easybourse.com/easybourse/secure/exportCsvValorisationTempsReel.html?siteLanguage=fr')
        files_before = set(os.listdir(self.download_dir))

        # Wait for file to download
        timeout = 30
        start_time = time.time()

        while time.time() - start_time < timeout:
            files_after = set(os.listdir(self.download_dir))
            new_files = files_after - files_before

            # Look for newly created CSV file
            csv_files = [f for f in new_files if f.endswith('.csv')]

            if csv_files:
                csv_filename = csv_files[0]
                logger.info(f"CSV file downloaded: {csv_filename}")
                return os.path.join(self.download_dir, csv_filename)

            time.sleep(1)

        return None

    def parse_csv_data(self, csv_path):
        """Parse CSV file and extract data with totals as columns"""
        try:
            # Read file with appropriate encoding
            with open(csv_path, 'r', encoding='cp1252') as f:
                content = f.read()

            lines = content.split('\n')

            # Extract valuation date (line 2)
            date_match = re.search(r'Valorisation au;(\d{2}/\d{2}/\d{4})', lines[2] if len(lines) > 2 else '')
            if date_match:
                date_str = date_match.group(1)
                valorisation_date = datetime.strptime(date_str, '%d/%m/%Y')
                logger.info(f"Valuation date: {valorisation_date.strftime('%d/%m/%Y')}")
            else:
                valorisation_date = datetime.now()
                logger.warning("Valuation date not found, using current date")

            # Extract totals from CSV (lines 7-12)
            totals_dict = {}

            logger.info("Extracting totals...")
            for i in range(7, 13):  # Lines 7 to 12
                if i < len(lines) and lines[i].strip():
                    parts = lines[i].split(';')
                    if len(parts) >= 2 and parts[0].strip():
                        label = parts[0].strip()
                        montant = parts[1].strip() if len(parts) > 1 else '0'

                        try:
                            # Convert amount
                            montant_float = float(montant.replace(',', '.').replace(' ', ''))

                            # Map labels to column names
                            if label == 'Total positions sous dossier':
                                totals_dict['Total positions sous dossier'] = montant_float
                            elif label == 'Solde espèces':
                                totals_dict['Solde espèces'] = montant_float
                            elif label == 'Valeur totale':
                                totals_dict['Valeur totale'] = montant_float

                            logger.info(f"  • {label}: {montant_float:,.2f}€")

                        except ValueError as e:
                            logger.warning(f"Unable to convert value for {label}: {montant}")

            # Find beginning of positions table
            header_index = -1
            for i, line in enumerate(lines):
                if 'Valeur;Code Isin;Place de cotation' in line:
                    header_index = i
                    break

            if header_index == -1:
                logger.error("Table headers not found")
                return None

            # Extract positions
            positions_lines = []
            for i in range(header_index, len(lines)):
                if lines[i].strip() and ';' in lines[i]:
                    positions_lines.append(lines[i])

            # Parse positions
            import io
            positions_content = '\n'.join(positions_lines)
            df = pd.read_csv(io.StringIO(positions_content), sep=';', decimal=',')

            # Clean columns
            df.columns = df.columns.str.strip()

            # Add date
            df['Date'] = valorisation_date

            # Convert numeric columns
            numeric_columns = ['Quantité', 'Cours', 'Prix moyen', 'Valorisation', '+/- value', 'Performance (%)',
                               'Poids']

            for col in numeric_columns:
                if col in df.columns:
                    def safe_convert(x):
                        if pd.isna(x):
                            return x
                        try:
                            x_str = str(x).strip()
                            x_str = x_str.replace(' ', '').replace(',', '.')
                            if col == 'Performance (%)' and '%' in x_str:
                                x_str = x_str.replace('%', '')
                            return float(x_str)
                        except:
                            return None

                    df[col] = df[col].apply(safe_convert)

            # IMPORTANT: Add totals as columns (same value for all rows)
            for col_name, value in totals_dict.items():
                df[col_name] = value

            logger.info(f"Total columns added: {list(totals_dict.keys())}")

            # Display summary
            logger.info(f"Data extracted: {len(df)} positions")
            if len(df) > 0:
                valorisation_totale = df['Valorisation'].sum()
                logger.info(f"Calculated positions valuation: {valorisation_totale:,.2f}€")
                if 'Total positions sous dossier' in totals_dict:
                    logger.info(f"Total positions (from CSV): {totals_dict['Total positions sous dossier']:,.2f}€")

            return df

        except Exception as e:
            logger.error(f"Error parsing CSV: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None

    def update_excel(self, df_new, excel_path='EasyBourse.xlsx'):
        """Update Excel file with new data including total columns"""
        try:
            logger.info(f"📁 Updating file: {excel_path}")

            # Total columns to manage
            TOTAL_COLUMNS = ['Valeur totale', 'Total positions sous dossier', 'Solde espèces']

            # Ensure total columns exist in df_new
            for col in TOTAL_COLUMNS:
                if col not in df_new.columns:
                    df_new[col] = None

            # Check if Excel file exists
            if os.path.exists(excel_path):
                # Read existing file
                df_existing = pd.read_excel(excel_path, sheet_name='Data')
                logger.info(f"📊 Existing file loaded: {len(df_existing)} rows")

                # Ensure total columns exist in df_existing
                for col in TOTAL_COLUMNS:
                    if col not in df_existing.columns:
                        df_existing[col] = None

                # Ensure dates are in correct format
                df_existing['Date'] = pd.to_datetime(df_existing['Date'])
                df_new['Date'] = pd.to_datetime(df_new['Date'])

                # Get unique dates from new CSV
                new_dates = df_new['Date'].unique()

                # Process each date from new CSV
                for date in new_dates:
                    # Extract data for this date
                    df_date_new = df_new[df_new['Date'] == date].copy()

                    # Get total values for this date (same for all rows)
                    totals_values = {}
                    for col in TOTAL_COLUMNS:
                        if col in df_date_new.columns and not df_date_new[col].isna().all():
                            totals_values[col] = df_date_new[col].iloc[0]  # Take first value

                    logger.info(f"📅 Processing date {date.strftime('%d/%m/%Y')}")
                    logger.info(f"   Totals: {totals_values}")

                    # Check if this date already exists in Excel file
                    if date in df_existing['Date'].values:
                        logger.info(f"✅ Date {date.strftime('%d/%m/%Y')} already exists in Excel")

                        # For each position of this date in new CSV
                        for idx, row in df_date_new.iterrows():
                            valeur = row['Valeur']

                            # Check if this stock already exists for this date
                            mask = (df_existing['Date'] == date) & (df_existing['Valeur'] == valeur)

                            if mask.any():
                                # Update existing row
                                logger.info(f"   🔄 Updating {valeur}")
                                update_idx = df_existing[mask].index[0]

                                # Update all columns
                                for col in df_new.columns:
                                    df_existing.at[update_idx, col] = row[col]
                            else:
                                # Add new position
                                logger.info(f"   ➕ Adding {valeur}")

                                # Find where to insert new row
                                date_mask = df_existing['Date'] == date
                                if date_mask.any():
                                    last_idx_for_date = df_existing[date_mask].index[-1]
                                    insert_idx = last_idx_for_date + 1
                                else:
                                    insert_idx = len(df_existing)

                                # Insert new row
                                df_existing = pd.concat([
                                    df_existing.iloc[:insert_idx],
                                    pd.DataFrame([row]),
                                    df_existing.iloc[insert_idx:]
                                ], ignore_index=True)

                        # Update total columns for all rows of this date
                        date_mask = df_existing['Date'] == date
                        for col, value in totals_values.items():
                            df_existing.loc[date_mask, col] = value
                            logger.info(f"   📊 Updating {col} = {value:,.2f} for all rows")

                    else:
                        # Date doesn't exist, add all rows
                        logger.info(f"🆕 New date {date.strftime('%d/%m/%Y')}, adding {len(df_date_new)} rows")

                        # Find where to insert according to chronological order
                        if len(df_existing) > 0:
                            dates_existing = df_existing['Date'].unique()
                            dates_existing_sorted = sorted(dates_existing)

                            insert_position = None
                            for i, existing_date in enumerate(dates_existing_sorted):
                                if date < existing_date:
                                    first_idx_of_date = df_existing[df_existing['Date'] == existing_date].index[0]
                                    insert_position = first_idx_of_date
                                    break

                            if insert_position is None:
                                insert_position = len(df_existing)

                            # Insert new data
                            df_existing = pd.concat([
                                df_existing.iloc[:insert_position],
                                df_date_new,
                                df_existing.iloc[insert_position:]
                            ], ignore_index=True)
                        else:
                            df_existing = df_date_new

                df_combined = df_existing

            else:
                # If file doesn't exist, use new data
                df_combined = df_new
                logger.info(f"🆕 Creating new Excel file with {len(df_new)} rows")

            # Sort by date then by value
            df_combined = df_combined.sort_values(['Date', 'Valeur'], ascending=[True, True])

            # Reset index
            df_combined = df_combined.reset_index(drop=True)

            # Clean df_combined: remove Unnamed columns
            df_combined = df_combined.loc[:, ~df_combined.columns.str.contains('Unnamed')]

            logger.info(f"📊 Total after update: {len(df_combined)} rows")

            # Save to Excel
            logger.info(f"💾 Saving to: {excel_path}")
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df_combined.to_excel(writer, sheet_name='Data', index=False)

                # Adjust column widths
                worksheet = writer.sheets['Data']
                for idx, column in enumerate(df_combined.columns):
                    column_length = max(
                        df_combined[column].astype(str).map(len).max(),
                        len(column)
                    )
                    # Handle columns beyond Z
                    if idx < 26:
                        col_letter = chr(65 + idx)
                    else:
                        col_letter = chr(65 + idx // 26 - 1) + chr(65 + idx % 26)
                    worksheet.column_dimensions[col_letter].width = min(column_length + 2, 50)

                # Format total columns with background color
                from openpyxl.styles import PatternFill, Font
                light_blue_fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
                bold_font = Font(bold=True)

                # Find total column indices
                for col_name in TOTAL_COLUMNS:
                    if col_name in df_combined.columns:
                        col_idx = df_combined.columns.get_loc(col_name) + 1  # +1 because Excel starts at 1

                        # Format header
                        header_cell = worksheet.cell(row=1, column=col_idx)
                        header_cell.fill = light_blue_fill
                        header_cell.font = bold_font

                        # Format data with light background
                        for row in range(2, len(df_combined) + 2):
                            cell = worksheet.cell(row=row, column=col_idx)
                            cell.fill = light_blue_fill

            logger.info(f"✅ Excel file updated: {excel_path}")
            return True

        except Exception as e:
            logger.error(f"❌ Error updating Excel file: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def backup_excel(self, excel_path):
        """Create a backup of the Excel file before modification"""
        if os.path.exists(excel_path):
            backup_dir = 'Save'
            os.makedirs(backup_dir, exist_ok=True)

            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_path = os.path.join(backup_dir, f'EasyBourse_backup_{timestamp}.xlsx')

            import shutil
            shutil.copy2(excel_path, backup_path)
            logger.info(f"Backup created: {backup_path}")

            # Clean old backups (keep only last 10)
            backups = sorted([f for f in os.listdir(backup_dir) if f.endswith('.xlsx')])
            if len(backups) > 10:
                for old_backup in backups[:-10]:
                    os.remove(os.path.join(backup_dir, old_backup))
                    logger.info(f"Old backup deleted: {old_backup}")

    def run(self, excel_path=None):
        """Execute complete process"""
        driver = None
        try:
            # Define Excel file path
            if excel_path is None:
                excel_path = 'EasyBourse.xlsx'

            # Create absolute path if necessary
            excel_path = os.path.abspath(excel_path)
            logger.info(f"Target Excel file: {excel_path}")

            # Configure driver
            driver = self.setup_driver()

            # Login
            if not self.login(driver):
                logger.error("Login failed")
                return False


            # Download CSV
            csv_path = self.download_valorisation_csv(driver)
            if not csv_path:
                logger.error("CSV download failed")
                logger.error("Check your internet connection or login credentials")
                return False

            # Parse data
            df = self.parse_csv_data(csv_path)
            if df is None:
                logger.error("CSV parsing failed")
                return False

            # Create backup if file exists
            if os.path.exists(excel_path):
                self.backup_excel(excel_path)

            # Update Excel - IMPORTANT: Pass excel_path as parameter!
            if self.update_excel(df, excel_path):  # <-- Fix here
                logger.info("✅ Process completed successfully!")

                # Optional: Delete downloaded CSV
                try:
                    os.remove(csv_path)
                    logger.info(f"Temporary CSV file deleted: {csv_path}")
                except Exception as e:
                    logger.warning(f"Unable to delete CSV: {e}")

                return True
            else:
                return False

        except Exception as e:
            logger.error(f"General error: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
        finally:
            if driver:
                driver.quit()


# Script usage
if __name__ == "__main__":
    # Replace with your credentials
    from logins import id, password

    USERNAME = id
    PASSWORD = password

    # Create and launch downloader
    downloader = EasyBourseValorisationDownloader(USERNAME, PASSWORD)
    downloader.run()