#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PASS Schedule PDF Generator
Automated tool to generate PDF schedules from PASS system with custom signature and message.
"""
import os
import sys
import time
import base64
import shutil
import io

# Sauvegarder la fonction print originale AVANT de la redéfinir
_original_print = print

def safe_print(message):
    """Print function that handles Unicode properly on all platforms"""
    try:
        _original_print(message)
        sys.stdout.flush()  # Force flush for subprocess
    except UnicodeEncodeError:
        # Fallback: encoder en UTF-8 puis décoder en cp1252 avec ignore
        try:
            encoded = message.encode('utf-8', errors='replace').decode('utf-8')
            _original_print(encoded)
            sys.stdout.flush()
        except:
            # Dernier recours : remplacer les emojis problématiques
            safe_message = message.replace('📅', '[DATE]').replace('⚠️', '[WARNING]').replace('→', '->')
            _original_print(safe_message)
            sys.stdout.flush()
from datetime import datetime, timedelta
from pathlib import Path

# Remplacer print par safe_print globalement
print = safe_print

# Third-party imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException

from reportlab.pdfgen import canvas
from reportlab.lib.colors import blue
from PyPDF2 import PdfReader, PdfWriter
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


# Custom Exceptions
class ScheduleGenerationError(Exception):
    """Exception personnalisée pour les erreurs de génération d'emploi du temps"""
    pass


class PDFProcessingError(Exception):
    """Exception personnalisée pour les erreurs de traitement PDF"""
    pass


class BrowserNavigationError(Exception):
    """Exception personnalisée pour les erreurs de navigation dans le navigateur"""
    pass

# Configuration constants
DEFAULT_TIMEOUT = 30
PDF_GENERATION_TIMEOUT = 15
PAGE_LOAD_TIMEOUT = 10
SIGNATURE_WIDTH = 80
SIGNATURE_HEIGHT = 40
TEXT_MARGIN = 30


class DateUtils:
    """Utilitaires pour la gestion des dates et semaines."""
    
    @staticmethod
    def get_monday_from_week_number(year: int, week_number: int) -> str:
        """Calcule la date du lundi d'une semaine donnée (format ISO)."""
        try:
            jan4 = datetime(year, 1, 4)
            monday_week1 = jan4 - timedelta(days=jan4.weekday())
            target_monday = monday_week1 + timedelta(weeks=week_number-1)
            return target_monday.strftime("%Y%m%d")
        except Exception as e:
            print(f"⚠️ Error calculating monday for week {week_number}: {e}")
            return datetime.now().strftime("%Y%m%d")
    
    @staticmethod
    def get_week_number_from_date(date_str: str) -> str:
        """Calcule le numéro de semaine (S37) à partir d'une date YYYYMMDD."""
        try:
            date_obj = datetime.strptime(date_str, "%Y%m%d")
            week_number = date_obj.isocalendar()[1]
            return f"S{week_number}"
        except Exception:
            return "S--"
    
    @staticmethod
    def get_target_date() -> str:
        """Détermine la date cible selon la priorité."""
        target_week = os.getenv('TARGET_WEEK')
        target_date = os.getenv('TARGET_DATE')
        weeks_offset = os.getenv('WEEKS_OFFSET')
        
        if target_week:
            try:
                week_num = int(target_week)
                current_year = datetime.now().year
                calculated_date = DateUtils.get_monday_from_week_number(current_year, week_num)
                print(f"📅 Using TARGET_WEEK={target_week} → {calculated_date}")
                return calculated_date
            except ValueError:
                print(f"⚠️ Invalid TARGET_WEEK value: {target_week}")
        
        if target_date:
            print(f"📅 Using TARGET_DATE: {target_date}")
            return target_date
        
        if weeks_offset:
            try:
                offset = int(weeks_offset)
                today = datetime.now()
                target = today + timedelta(weeks=offset)
                calculated_date = target.strftime("%Y%m%d")
                print(f"📅 Using WEEKS_OFFSET={weeks_offset} → {calculated_date}")
                return calculated_date
            except ValueError:
                print(f"⚠️ Invalid WEEKS_OFFSET value: {weeks_offset}")
        
        current_date = datetime.now().strftime("%Y%m%d")
        print(f"📅 Using current date: {current_date}")
        return current_date


class FileUtils:
    """Utilitaires pour la gestion des fichiers."""
    
    @staticmethod
    def clean_filename_for_windows(filename: str) -> str:
        """Nettoie un nom de fichier pour Windows."""
        forbidden_chars = '<>:"|?*\\/'
        cleaned = filename
        for char in forbidden_chars:
            cleaned = cleaned.replace(char, '-')
        return cleaned.strip()
    
    @staticmethod
    def create_pdf_filename(target_date: str = None) -> str:
        """Crée le nom de fichier PDF."""
        nom_prenom = os.getenv('NOM_PRENOM', 'Nom Prénom')
        promo = os.getenv('PROMO', 'PROMO')
        
        if target_date:
            semaine = DateUtils.get_week_number_from_date(target_date)
        else:
            semaine = DateUtils.get_week_number_from_date(datetime.now().strftime("%Y%m%d"))
        
        filename = f"{nom_prenom} – {promo} – {semaine}.pdf"
        return FileUtils.clean_filename_for_windows(filename)
    
    @staticmethod
    def ensure_directory_exists(directory: str) -> None:
        """Crée un répertoire s'il n'existe pas."""
        Path(directory).mkdir(parents=True, exist_ok=True)
    
    @staticmethod
    def safe_rename_pdf(temp_path: str, target_filename: str, save_folder: str) -> str:
        """Renomme le PDF de manière sécurisée."""
        try:
            final_path = Path(save_folder) / target_filename
            shutil.move(temp_path, final_path)
            print(f"✅ PDF renamed to: {target_filename}")
            return str(final_path)
        except Exception as e:
            print(f"⚠️ Could not rename to '{target_filename}': {e}")
            fallback_name = target_filename.replace("–", "-").replace(" ", "_")
            try:
                fallback_path = Path(save_folder) / fallback_name
                shutil.move(temp_path, fallback_path)
                print(f"✅ PDF saved with fallback name: {fallback_name}")
                return str(fallback_path)
            except Exception as e2:
                print(f"❌ Rename failed: {e2}")
                return temp_path


def get_week_date(weeks_offset=0):
    """
    Calcule la date d'une semaine spécifique
    weeks_offset: nombre de semaines à partir de maintenant
    - 0 = semaine actuelle
    - 1 = semaine prochaine  
    - -1 = semaine dernière
    """
    today = datetime.now()
    target_date = today + timedelta(weeks=weeks_offset)
    return target_date.strftime("%Y%m%d")


def add_message_to_pdf(input_pdf_path, output_pdf_path, message):
    """Ajoute un message personnalisé et une signature sur le PDF avec gestion d'erreurs robuste"""
    try:
        print(f"✏️ Adding custom message and signature to PDF: '{message}'")
        
        # Validation des paramètres d'entrée
        if not input_pdf_path or not os.path.exists(input_pdf_path):
            raise PDFProcessingError(f"Input PDF file not found: {input_pdf_path}")
        
        if not message or not message.strip():
            raise PDFProcessingError("Message cannot be empty")
        
        # Lire le PDF original
        try:
            reader = PdfReader(input_pdf_path)
        except Exception as e:
            raise PDFProcessingError(f"Failed to read input PDF: {e}")
        
        if len(reader.pages) == 0:
            raise PDFProcessingError("Input PDF has no pages")
        
        writer = PdfWriter()
        
        # Traiter chaque page (normalement juste la première)
        for page_num, page in enumerate(reader.pages):
            try:
                # Créer un overlay avec le message et la signature
                packet = io.BytesIO()
                
                # Obtenir les dimensions de la page
                page_width = float(page.mediabox.width)
                page_height = float(page.mediabox.height)
                
                print(f"🔍 Page {page_num + 1} dimensions: width={page_width}, height={page_height}")
                
                # Créer un canvas pour le texte overlay
                can = canvas.Canvas(packet, pagesize=(page_width, page_height))
                
                # Configuration du texte - position en bas à gauche
                can.setFont("Helvetica-Bold", 12)
                can.setFillColor(blue)
                
                # Positionner le message en bas à gauche (avec marge de sécurité)
                text_x = 30  # 30 points du bord gauche
                text_y = 30  # 30 points du bas
                
                print(f"📝 Adding text at position: x={text_x}, y={text_y}")
                can.drawString(text_x, text_y, message)
                
                # Ajouter la signature en bas à droite
                signature_file = os.getenv('SIGNATURE_FILE', 'signature.png')
                signature_path = os.path.abspath(signature_file)
                
                if signature_file and os.path.exists(signature_path):
                    try:
                        print(f"📝 Adding signature from: {signature_path}")
                        
                        # Dimensions de la signature (plus petites pour être sûr)
                        signature_width = 80  # largeur en points
                        signature_height = 40  # hauteur en points
                        
                        # Position en bas à droite (avec marges de sécurité)
                        sig_x = page_width - signature_width - 30  # 30 points du bord droit
                        sig_y = 30  # 30 points du bas
                        
                        print(f"🖼️ Adding signature at position: x={sig_x}, y={sig_y}, w={signature_width}, h={signature_height}")
                        
                        # Ajouter l'image de signature
                        can.drawImage(signature_path, sig_x, sig_y, 
                                    width=signature_width, height=signature_height, 
                                    mask='auto')  # Support de la transparence
                        print("✅ Signature added successfully")
                    except Exception as sig_error:
                        print(f"⚠️ Warning: Could not add signature: {sig_error}")
                        # Continue sans signature plutôt que d'échouer complètement
                else:
                    print(f"⚠️ Signature file not found or not configured: {signature_path}")
                
                can.save()
                
                # Revenir au début du buffer
                packet.seek(0)
                
                # Créer un nouveau PDF à partir du buffer
                try:
                    overlay_pdf = PdfReader(packet)
                    # Fusionner avec la page originale
                    page.merge_page(overlay_pdf.pages[0])
                except Exception as merge_error:
                    print(f"⚠️ Warning: Could not merge overlay on page {page_num + 1}: {merge_error}")
                    # Ajouter la page originale sans overlay
                
                writer.add_page(page)
                
            except Exception as page_error:
                print(f"⚠️ Warning: Error processing page {page_num + 1}: {page_error}")
                # Ajouter la page originale en cas d'erreur
                writer.add_page(page)
        
        # Sauvegarder le PDF modifié
        try:
            # S'assurer que le dossier de destination existe
            output_dir = os.path.dirname(output_pdf_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            with open(output_pdf_path, 'wb') as output_file:
                writer.write(output_file)
        except Exception as save_error:
            raise PDFProcessingError(f"Failed to save output PDF: {save_error}")
        
        print(f"✅ PDF with custom message and signature saved to: {output_pdf_path}")
        return True
        
    except PDFProcessingError:
        # Re-raise les erreurs PDF spécifiques
        raise
    except Exception as e:
        print(f"❌ Unexpected error adding message/signature to PDF: {e}")
        import traceback
        traceback.print_exc()
        
        # En cas d'erreur, essayer de copier le fichier original
        try:
            if os.path.exists(input_pdf_path):
                shutil.copy2(input_pdf_path, output_pdf_path)
                print(f"📋 Copied original PDF to output location as fallback")
        except Exception as copy_error:
            print(f"❌ Failed to copy original PDF as fallback: {copy_error}")
            
        raise PDFProcessingError(f"Failed to process PDF: {e}")


def login(driver, wait, username, password):
    """Se connecte à PASS en utilisant SSO avec gestion d'erreurs robuste"""
    try:
        print("Connecting to PASS...")
        
        # Validation des paramètres d'entrée
        if not username or not password:
            raise BrowserNavigationError("Username and password are required")
        
        # Navigation vers PASS
        try:
            driver.get('https://pass.imt-atlantique.fr')
            print("✅ Successfully navigated to PASS")
        except Exception as e:
            raise BrowserNavigationError(f"Failed to navigate to PASS: {e}")
        
        # Attendre que la page soit complètement chargée
        time.sleep(2)
        
        # Cliquer sur le bouton SSO
        try:
            print("🔍 Looking for SSO button...")
            sso_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='SSO']]")))
            sso_button.click()
            print("✅ SSO button clicked")
            time.sleep(3)  # Attendre la redirection
        except TimeoutException:
            raise BrowserNavigationError("SSO button not found or not clickable")
        except Exception as e:
            raise BrowserNavigationError(f"Failed to click SSO button: {e}")

        # Saisir les identifiants
        try:
            print("🔍 Looking for login form...")
            username_form = wait.until(EC.visibility_of_element_located((By.ID, 'username')))
            password_form = wait.until(EC.visibility_of_element_located((By.ID, 'password')))
            print("✅ Login form found")
        except TimeoutException:
            raise BrowserNavigationError("Login form not found or not visible")
        except Exception as e:
            raise BrowserNavigationError(f"Failed to locate login form: {e}")
        
        # Saisir le nom d'utilisateur avec délai
        try:
            print("📝 Entering username...")
            username_form.clear()
            username_form.send_keys(username)
            print("✅ Username entered")
        except Exception as e:
            raise BrowserNavigationError(f"Failed to enter username: {e}")
        
        # Saisir le mot de passe avec délai
        try:
            print("📝 Entering password...")
            password_form.clear()
            password_form.send_keys(password)
            print("✅ Password entered")
        except Exception as e:
            raise BrowserNavigationError(f"Failed to enter password: {e}")

        # Soumettre le formulaire de connexion
        try:
            print("🔍 Looking for submit button...")
            submit_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-submit')))
            submit_button.click()
            print("✅ Login form submitted")
            time.sleep(4)  # Attendre la validation
        except TimeoutException:
            raise BrowserNavigationError("Submit button not found or not clickable")
        except Exception as e:
            raise BrowserNavigationError(f"Failed to submit login form: {e}")

        # Confirmer la connexion
        try:
            print("🔍 Looking for confirmation button...")
            confirm_button = wait.until(EC.element_to_be_clickable((By.NAME, '_eventId_proceed')))
            confirm_button.click()
            print("✅ Login confirmed")
            time.sleep(3)  # Attendre la redirection finale
        except TimeoutException:
            raise BrowserNavigationError("Confirmation button not found or not clickable")
        except Exception as e:
            raise BrowserNavigationError(f"Failed to confirm login: {e}")
        
        # Vérifier que la connexion a réussi
        try:
            # Attendre d'être redirigé vers la page principale
            current_url = driver.current_url
            if "pass.imt-atlantique.fr" not in current_url:
                raise BrowserNavigationError(f"Unexpected redirect after login: {current_url}")
            print("✅ Login completed successfully")
        except Exception as e:
            print(f"⚠️ Warning: Could not verify login success: {e}")
            # Continue anyway, the login might still have worked
    
    except BrowserNavigationError:
        # Re-raise navigation errors
        raise
    except Exception as e:
        raise BrowserNavigationError(f"Unexpected error during login: {e}")


def navigate_to_schedule(driver, wait):
    """Navigue vers l'emploi du temps dans PASS avec gestion d'erreurs robuste"""
    try:
        print("Navigating to schedule...")
        
        # Attendre un peu avant de naviguer
        time.sleep(0.5)
        
        # Naviguer directement vers la page d'emploi du temps
        try:
            driver.get('https://pass.imt-atlantique.fr/OpDotNet/Noyau/Default.aspx?')
            print("✅ Successfully navigated to schedule page")
        except Exception as e:
            raise BrowserNavigationError(f"Failed to navigate to schedule page: {e}")
        
        # Attendre que la page soit complètement chargée
        try:
            wait.until(lambda driver: driver.execute_script('return document.readyState=="complete"'))
            print("✅ Page fully loaded")
        except TimeoutException:
            print("⚠️ Warning: Page load timeout, continuing anyway")
        except Exception as e:
            print(f"⚠️ Warning: Error checking page load state: {e}")
        
        # Attendre que jQuery soit chargé (souvent utilisé dans ces applications)
        print("🔍 Waiting for scripts to load...")
        try:
            wait.until(lambda driver: driver.execute_script('return typeof jQuery !== "undefined"'))
            print("✅ jQuery loaded")
        except TimeoutException:
            print("⚠️ jQuery not found or timeout, continuing...")
        except Exception as e:
            print(f"⚠️ Error checking jQuery: {e}")
        
        # Attendre un peu plus pour que tous les éléments soient chargés
        time.sleep(2)
        
        print("🔍 Debug: Checking frames structure...")
        try:
            # Lister tous les frames disponibles
            frames = driver.find_elements(By.TAG_NAME, "frame")
            print(f"Found {len(frames)} frames:")
            
            for i, frame in enumerate(frames):
                try:
                    name = frame.get_attribute("name") or "unnamed"
                    src = frame.get_attribute("src") or "no-src"
                    print(f"  Frame {i+1}: name='{name}' src='{src}'")
                except Exception as frame_error:
                    print(f"  Frame {i+1}: Error reading attributes: {frame_error}")
            
            # Chercher le frame content qui contient l'emploi du temps
            content_frame = None
            for frame in frames:
                try:
                    if frame.get_attribute("name") == "content":
                        content_frame = frame
                        break
                except Exception as e:
                    print(f"⚠️ Warning: Error checking frame name: {e}")
                    continue
            
            if content_frame:
                try:
                    print("✅ Found 'content' frame, switching to it...")
                    driver.switch_to.frame("content")
                    # Attendre que le contenu du frame soit chargé
                    time.sleep(5)
                    print("✅ Successfully switched to content frame")
                except Exception as switch_error:
                    print(f"⚠️ Warning: Error switching to content frame: {switch_error}")
                    driver.switch_to.default_content()  # Retour au contexte principal
            else:
                print("⚠️ No 'content' frame found, staying in main context")
                
        except Exception as e:
            print(f"⚠️ Warning: Error checking frames structure: {e}")
            # Continue anyway, maybe the page structure is different
        
        print("✅ Successfully navigated to schedule page.")
    
    except BrowserNavigationError:
        # Re-raise navigation errors
        raise
    except Exception as e:
        raise BrowserNavigationError(f"Unexpected error navigating to schedule: {e}")


def generate_schedule_pdf(driver, save_folder="pdfs", target_date=None):
    """Génère un PDF de l'emploi du temps en ouvrant directement l'URL de l'iframe Agenda.asp"""
    print("Generating schedule PDF... [VERSION: Direct iframe URL navigation]")
    
    # Créer le dossier s'il n'existe pas
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)
        print(f"Created folder: {save_folder}")
    
    # Attendre que la page soit bien chargée
    print("Waiting for page to fully load...")
    time.sleep(3)
    
    # S'assurer qu'on est dans le bon frame (content) pour trouver l'iframe
    try:
        print("🔍 Ensuring we're in the content frame to find iframe URL...")
        driver.switch_to.default_content()  # Retour au contexte principal
        driver.switch_to.frame("content")   # Puis switch vers content
        print("✅ Switched to content frame")
        
        # Chercher l'iframe qui contient l'agenda et extraire son URL
        print("🔍 Looking for iframe URL in content frame...")
        try:
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            print(f"Found {len(iframes)} iframes in content frame")
            
            agenda_iframe_url = None
            for i, iframe in enumerate(iframes):
                src = iframe.get_attribute("src")
                print(f"  Iframe {i+1}: src='{src}'")
                if src and "Agenda.asp" in src:
                    agenda_iframe_url = src
                    print(f"✅ Found agenda iframe URL: {src}")
                    break
            
            if agenda_iframe_url:
                print("🌐 Opening iframe URL in NEW TAB...")
                # Ouvrir un nouvel onglet et naviguer vers l'URL de l'iframe
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[-1])  # Aller au nouvel onglet
                driver.get(agenda_iframe_url)
                
                # Attendre que la page de l'agenda soit complètement chargée
                print("⏳ Waiting for agenda page to load in new tab...")
                time.sleep(8)
                
                # Vérifier que la page est bien chargée
                page_length = len(driver.page_source)
                links_count = len(driver.find_elements(By.TAG_NAME, "a"))
                print(f"✅ Agenda page loaded: page_length={page_length}, links={links_count}")
                
                # Naviguer vers la date cible si spécifiée
                if target_date:
                    print(f"📅 Navigating to week of {target_date} in iframe...")
                    try:
                        driver.execute_script(f"NavDat('{target_date}');")
                        time.sleep(3)
                        print(f"✅ Successfully navigated to week of {target_date}")
                    except Exception as e:
                        print(f"⚠️ Could not navigate to date {target_date}: {e}")
                
            else:
                print("❌ No agenda iframe found, staying in content frame")
                
        except Exception as e:
            print(f"❌ Error looking for iframe: {e}")
        
        time.sleep(3)  # Délai supplémentaire pour être sûr
    except Exception as e:
        print(f"❌ Error in iframe navigation: {e}")
        # Si on ne peut pas trouver l'iframe, on continue quand même
    
    # Générer un nom de fichier propre : "GUERRY Roman – FIPA3R – S38.pdf"
    pdf_filename = FileUtils.create_pdf_filename(target_date)
    expected_pdf_path = os.path.join(save_folder, pdf_filename)
    
    # Si le fichier existe déjà, il sera overwrité
    print(f"📄 PDF filename: {pdf_filename}")
    
    try:
        # Étape 1: Chercher le bouton d'impression de PASS AVANT tout autre chose
        print("🔍 Looking for PASS print button...")
        
        # Essayer plusieurs sélecteurs pour trouver le bouton d'impression de PASS
        selectors = [
            "//a[@onclick and contains(@onclick, 'Imprimer')]",
            "//a[contains(@onclick, 'Imprimer()')]", 
            "//img[@src='/dataop/visuel/icones/16x16/Imprimer.gif']/..",
            "//img[contains(@src, 'Imprimer.gif')]/..",
            "//a[@title='Imprimer cette visualisation']",
            "//img[contains(@src, 'print')]/..",
            "//a[contains(@onclick, 'print')]",
            "//a[contains(text(), 'Imprimer')]"
        ]
        
        print_button = None
        for selector in selectors:
            try:
                print(f"  Trying selector: {selector}")
                print_button = driver.find_element(By.XPATH, selector)
                onclick_attr = print_button.get_attribute("onclick")
                print(f"✅ Found print button with selector: {selector}")
                print(f"   onclick attribute: {onclick_attr}")
                break
            except Exception as e:
                print(f"❌ Selector failed: {selector}")
                continue
        
        if print_button is None:
            print("❌ ERROR: Could not find PASS print button")
            return None
        
        # Étape 2: Préparer l'interception de window.print() AVANT de cliquer
        print("�️ Preparing to intercept window.print() and prevent Windows dialog...")
        
        # Remplacer window.print() par Chrome DevTools Protocol
        driver.execute_script("""
            // Sauvegarder la fonction originale
            window.originalPrint = window.print;
            
            // Remplacer window.print() par notre version qui utilise CDP
            window.print = function() {
                console.log('window.print() intercepted - will use Chrome DevTools Protocol instead');
                
                // Ajouter du CSS d'impression
                var printStyle = document.createElement('style');
                printStyle.innerHTML = `
                    @media print {
                        body { margin: 0; padding: 5px; font-size: 12px; }
                        * { -webkit-print-color-adjust: exact !important; }
                        .no-print, .noprint { display: none !important; }
                        table { 
                            page-break-inside: avoid; 
                            border-collapse: collapse;
                            width: 100% !important;
                        }
                    }
                    @page { 
                        size: A4 landscape; 
                        margin: 0.5cm; 
                    }
                `;
                document.head.appendChild(printStyle);
                
                // Déclencher un événement personnalisé pour notre script Python
                window.dispatchEvent(new CustomEvent('passPrintRequested'));
                
                // NE PAS appeler la fonction print originale pour éviter le dialogue
                return false;
            };
            
            console.log('window.print() successfully overridden');
        """)
        
        print("✅ window.print() override installed")
        time.sleep(1)
        
        # Étape 3: Cliquer sur le bouton Imprimer de PASS
        print("🖱️ Clicking PASS print button...")
        print_button.click()
        
        # Étape 4: Attendre un court délai puis utiliser Chrome DevTools Protocol
        print("⏳ Waiting for print request, then using Chrome DevTools Protocol...")
        time.sleep(3)
        
        # Configuration d'impression pour Microsoft Print to PDF équivalente
        print_settings = {
            'landscape': True,
            'displayHeaderFooter': False,
            'printBackground': True,
            'scale': 0.8,  # Réduire légèrement pour mieux s'adapter
            'paperWidth': 11.69,  # A4 width in inches (landscape)
            'paperHeight': 8.27,  # A4 height in inches (landscape)
            'marginTop': 0.4,
            'marginBottom': 0.4,
            'marginLeft': 0.4,
            'marginRight': 0.4,
            'pageRanges': '1',  # Imprimer seulement la première page
            'headerTemplate': '',
            'footerTemplate': '',
            'preferCSSPageSize': False,
            'generateTaggedPDF': False,
            'generateDocumentOutline': False
        }
        
        print("📄 Generating PDF with Chrome DevTools Protocol...")
        result = driver.execute_cdp_cmd('Page.printToPDF', print_settings)
        
        if 'data' in result:
            print("✅ PDF generated successfully!")
            
            # Décoder le PDF base64 et le sauvegarder
            pdf_data = base64.b64decode(result['data'])
            
            # Sauvegarder le PDF temporaire avec un nom simple d'abord
            temp_pdf_path = expected_pdf_path.replace('.pdf', '_temp.pdf')
            full_temp_path = os.path.abspath(temp_pdf_path)
            with open(full_temp_path, 'wb') as f:
                f.write(pdf_data)
            
            print(f"📁 Temporary PDF saved to: {full_temp_path}")
            
            # Ajouter le message personnalisé au PDF avec nom temporaire
            temp_final_path = expected_pdf_path.replace('.pdf', '_with_message.pdf')
            pdf_message = os.getenv('PDF_MESSAGE', 'Emploi du temps généré automatiquement')
            
            if add_message_to_pdf(full_temp_path, temp_final_path, pdf_message):
                # Supprimer le fichier temporaire sans message
                os.remove(full_temp_path)
                
                # Renommer vers le nom final souhaité
                final_path = FileUtils.safe_rename_pdf(temp_final_path, pdf_filename, save_folder)
                print(f"📁 Final PDF saved to: {final_path}")
            else:
                # En cas d'erreur, renommer le fichier temporaire directement
                final_path = FileUtils.safe_rename_pdf(full_temp_path, pdf_filename, save_folder)
                print(f"📁 PDF saved to: {final_path}")
            
            return final_path
        else:
            print("❌ CDP method failed, no data returned")
            return None
            
    except Exception as e:
        print(f"❌ Error during PDF generation: {e}")
        return None


def create_optimized_chrome_options(save_folder: str) -> Options:
    """
    Crée une configuration Chrome optimisée pour l'automatisation PDF.
    
    Args:
        save_folder: Dossier de sauvegarde des PDF
        
    Returns:
        Options Chrome configurées
    """
    options = Options()
    
    # Performance et stabilité
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-plugins')
    options.add_argument('--disable-background-timer-throttling')
    options.add_argument('--disable-backgrounding-occluded-windows')
    options.add_argument('--disable-renderer-backgrounding')
    
    # Impression automatique
    options.add_argument('--kiosk-printing')
    options.add_argument('--disable-print-preview')
    options.add_argument('--disable-popup-blocking')
    options.add_argument('--use-fake-ui-for-media-stream')
    
    # Préférences optimisées
    prefs = {
        "printing.print_preview_sticky_settings.appState": {
            "recentDestinations": [{
                "id": "Microsoft Print to PDF",
                "origin": "local",
                "account": ""
            }],
            "selectedDestinationId": "Microsoft Print to PDF",
            "version": 2,
            "isHeaderFooterEnabled": False,
            "isLandscapeEnabled": True,
            "isCssBackgroundEnabled": True,
            "marginsType": 1,
            "scaling": 100,
            "shouldPrintBackgrounds": True,
            "shouldPrintSelectionOnly": False
        },
        "savefile.default_directory": os.path.abspath(save_folder),
        "download.default_directory": os.path.abspath(save_folder),
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": False,
        "printing.use_system_print_dialog": False,
        "printing.print_preview_disabled": True,
        "profile.default_content_setting_values.notifications": 2
    }
    options.add_experimental_option("prefs", prefs)
    
    return options


def main(username, password, nom_prenom, promo, target_date, save_folder, debug_mode):
    """Fonction principale avec gestion d'erreurs complète."""
    print("🚀 Starting schedule PDF generation process...")
    
    driver = None
    
    try:
        # Validation de la configuration
        missing_vars = []
        if not username:
            missing_vars.append("USERNAME")
        if not password:
            missing_vars.append("PASSWORD")
        if not nom_prenom:
            missing_vars.append("NOM_PRENOM")
        if not promo:
            missing_vars.append("PROMO")
        
        if missing_vars:
            raise ScheduleGenerationError(f"Missing environment variables: {', '.join(missing_vars)}")



        # Ensure localhost is not proxied (for WSL2 and similar environments)
        no_proxy = os.environ.get("NO_PROXY", "")
        needed = ["127.0.0.1", "localhost", "::1"]
        for host in needed:
            if host not in no_proxy:
                no_proxy = f"{no_proxy},{host}" if no_proxy else host
        os.environ["NO_PROXY"] = no_proxy
        os.environ["no_proxy"] = no_proxy


        print(f"✅ Configuration validated for {nom_prenom} ({promo})")
    

        # Créer le dossier de sauvegarde
        try:
            FileUtils.ensure_directory_exists(save_folder)
            print(f"✅ Save folder ready: {os.path.abspath(save_folder)}")
        except Exception as e:
            raise ScheduleGenerationError(f"Failed to create save folder: {e}")
        
        # Configuration Chrome optimisée
        try:
            options = create_optimized_chrome_options(save_folder)
            print("✅ Chrome options configured")
        except Exception as e:
            raise ScheduleGenerationError(f"Failed to configure Chrome options: {e}")
        
        # Démarrer le navigateur
        try:
            driver = webdriver.Chrome(options=options)
            wait = WebDriverWait(driver, 60)
            print("✅ Chrome browser started")
        except Exception as e:
            raise ScheduleGenerationError(f"Failed to start Chrome browser: {e}")
        
        # Se connecter à PASS
        print("🔐 Starting login process...")
        try:
            login(driver, wait, username, password)
            print("✅ Login successful")
        except BrowserNavigationError as e:
            raise ScheduleGenerationError(f"Login failed: {e}")
        
        # Naviguer vers l'emploi du temps
        print("🧭 Navigating to schedule...")
        try:
            navigate_to_schedule(driver, wait)
            print("✅ Navigation successful")
        except BrowserNavigationError as e:
            raise ScheduleGenerationError(f"Navigation failed: {e}")
        
        # Générer le PDF en utilisant la fonction Imprimer() de PASS
        print("📄 Generating PDF...")
        try:
            pdf_path = generate_schedule_pdf(driver, save_folder, target_date)
            print(f"✅ PDF generated: {pdf_path}")
        except (ScheduleGenerationError, PDFProcessingError) as e:
            raise ScheduleGenerationError(f"PDF generation failed: {e}")
        
        # Message de confirmation final
        date_info = f" (semaine du {target_date})" if target_date else ""
        print(f"🎉 PDF de l'emploi du temps généré avec succès{date_info}!")
        print(f"📁 Dossier: {os.path.abspath(save_folder)}")
        if pdf_path:
            print(f"📄 Fichier: {os.path.basename(pdf_path)}")

    except ScheduleGenerationError as e:
        print(f"❌ Erreur de génération: {e}")
        return 1
    except PDFProcessingError as e:
        print(f"❌ Erreur de traitement PDF: {e}")
        return 1
    except BrowserNavigationError as e:
        print(f"❌ Erreur de navigation: {e}")
        return 1
    except Exception as e:
        print(f"❌ Erreur inattendue: {e}")
        import traceback
        traceback.print_exc()
        return 1
    finally:
        if driver:
            try:
                if debug_mode:
                    print("🔍 DEBUG MODE: Navigateur laissé ouvert pour inspection")
                    print("Appuyez sur Entrée pour fermer le navigateur...")
                    input()
                driver.quit()
                print("✅ Browser closed successfully")
            except Exception as cleanup_error:
                print(f"⚠️ Warning: Error closing browser: {cleanup_error}")
    
    return 0


if __name__ == "__main__":
    # Charger les variables d'environnement
    load_dotenv()
    
    # Configuration des variables
    username = os.getenv('IMT_USERNAME')
    password = os.getenv('IMT_PASSWORD')
    nom_prenom = os.getenv('NOM_PRENOM')
    promo = os.getenv('PROMO')

    # Configuration de la semaine à capturer
    # TARGET_WEEK : numéro de semaine ISO (ex: '38' pour semaine 38)
    # TARGET_DATE : date spécifique au format YYYYMMDD (ex: '20250915') 
    # WEEKS_OFFSET : nombre de semaines à partir de maintenant (ex: '0', '1', '-1')
    target_date = DateUtils.get_target_date()  # Utilise la méthode moderne des utilitaires
    save_folder = os.getenv('SAVE_FOLDER', 'pdfs')  # Dossier de sauvegarde pour PDFs
    debug_mode = os.getenv('DEBUG_MODE', 'false').lower() == 'true'  # Mode debug

    print(f"Configuration: username={username}, target_date={target_date}, save_folder={save_folder}, debug_mode={debug_mode}")
    
    try:
        # Message de début
        print("📄 Starting schedule PDF generation...")
        
        # Exécuter la génération PDF
        main(username, password, nom_prenom, promo, target_date, save_folder, debug_mode)
        
        # Message de fin
        print("✅ Schedule PDF generation completed")
        
        # En mode debug, proposer de relancer
        if debug_mode:
            print("🔍 DEBUG MODE: Voulez-vous relancer ? (Entrée = Oui, Ctrl+C = Non)")
            try:
                input()
                print("Relancement...")
                # Note: En mode debug, on pourrait ajouter une boucle ici si souhaité
            except KeyboardInterrupt:
                print("🛑 Script arrêté par l'utilisateur.")
        else:
            print("✅ Script terminé avec succès.")
                
    except KeyboardInterrupt:
        print("🛑 Script arrêté par l'utilisateur.")