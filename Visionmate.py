"""
automation_gui_tkinter.py
Tkinter GUI wrapper for Google Maps + email scraping automation.

Run:
    python automation_gui_tkinter.py
"""

import os
import time
import re
import shutil
import threading
import platform
import requests
import pandas as pd
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# ---------- Sleep prevention (Windows) ----------
def prevent_sleep_windows():
    try:
        if platform.system() != "Windows":
            return None
        import ctypes
        ES_CONTINUOUS = 0x80000000
        ES_SYSTEM_REQUIRED = 0x00000001
        ES_DISPLAY_REQUIRED = 0x00000002
        flags = ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        ctypes.windll.kernel32.SetThreadExecutionState(flags)
        return True
    except Exception as e:
        print("Could not set Windows sleep prevention:", e)
        return None

def clear_sleep_windows():
    try:
        if platform.system() != "Windows":
            return None
        import ctypes
        ES_CONTINUOUS = 0x80000000
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        return True
    except Exception as e:
        print("Could not clear Windows sleep prevention:", e)
        return None

# ---------- Selenium setup ----------
def setup_driver(headless=False):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

# ---------- Business detail scraping ----------
def get_business_details(driver):
    details = {}
    try:
        details["Name"] = driver.find_element(By.XPATH, '//h1[contains(@class,"DUwDvf")]').text
    except:
        details["Name"] = "Not Available"

    try:
        details["Address"] = driver.find_element(By.XPATH, '//button[contains(@aria-label,"Address")]').text
    except:
        details["Address"] = "Not Available"

    try:
        details["Phone"] = driver.find_element(By.XPATH, '//button[contains(@aria-label,"Phone")]').text
    except:
        details["Phone"] = "Not Available"

    try:
        details["Website"] = driver.find_element(By.XPATH, '//a[contains(@aria-label,"Website")]').get_attribute("href")
    except:
        details["Website"] = "Not Available"

    try:
        details["Email"] = driver.find_element(By.XPATH, '//a[starts-with(@href,"mailto:")]').get_attribute("href").replace("mailto:", "")
    except:
        details["Email"] = "Not Available"

    return details

def collect_all_listings(driver, gui, base_query):
    listings_data = []
    url = f"https://www.google.com/maps/search/{requests.utils.quote(base_query)}/"
    driver.get(url)
    time.sleep(5)

    try:
        scrollable_div = driver.find_element(By.XPATH, '//div[contains(@aria-label,"Results for")]')
    except:
        try:
            scrollable_div = driver.find_element(By.XPATH, '//div[contains(@class,"m6QErb") and contains(@aria-label,"Results")]')
        except:
            gui.log("Could not locate results panel on the page.")
            return listings_data

    last_height = 0
    stable_scrolls = 0
    collected = 0

    while True:
        listings = driver.find_elements(By.XPATH, '//div[contains(@class,"Nv2PK")]')
        gui.update_status(f"Found {len(listings)} listings so far...")

        for i in range(len(listings_data), len(listings)):
            try:
                driver.execute_script("arguments[0].scrollIntoView();", listings[i])
                time.sleep(1)
                listings[i].click()
                time.sleep(4)

                details = get_business_details(driver)
                listings_data.append(details)
                collected += 1

                gui.log(f"Collected: {details.get('Name')}")
                gui.update_status(f"Collected {collected} listings...")

                try:
                    driver.find_element(By.XPATH, '//button[@aria-label="Back"]').click()
                    time.sleep(1.5)
                except:
                    pass

                listings = driver.find_elements(By.XPATH, '//div[contains(@class,"Nv2PK")]')
            except Exception as e:
                gui.log(f"Error collecting listing {i}: {e}")
                continue

        try:
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scrollable_div)
        except Exception:
            driver.execute_script("window.scrollBy(0, 1000);")
        time.sleep(2)

        try:
            new_height = driver.execute_script("return arguments[0].scrollHeight", scrollable_div)
        except:
            new_height = driver.execute_script("return document.body.scrollHeight")

        if new_height == last_height:
            stable_scrolls += 1
        else:
            stable_scrolls = 0

        if stable_scrolls >= 3:
            gui.log("All listings loaded.")
            break

        last_height = new_height

    return listings_data

# ---------- Email scraping ----------
def find_emails_from_website(url):
    try:
        if not url or not isinstance(url, str):
            return None
        if url.strip().lower() == "not available":
            return None
        if not url.startswith("http"):
            url = "http://" + url

        resp = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
        text = resp.text
        mailtos = [a["href"].replace("mailto:", "") for a in BeautifulSoup(text, "html.parser").find_all("a", href=True) if a["href"].startswith("mailto:")]
        text_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-z]{2,}", text)
        all_emails = list(set(mailtos + text_emails))

        if not all_emails:
            soup = BeautifulSoup(text, "html.parser")
            contact_links = [a["href"] for a in soup.find_all("a", href=True) if "contact" in a["href"].lower()]
            if contact_links:
                contact_url = contact_links[0]
                if not contact_url.startswith("http"):
                    contact_url = requests.compat.urljoin(url, contact_url)
                try:
                    creq = requests.get(contact_url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
                    contact_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-z]{2,}", creq.text)
                    all_emails = list(set(contact_emails))
                except:
                    pass

        return ", ".join(all_emails) if all_emails else None
    except:
        return None

# ---------- Worker ----------
def worker_scrape(base_query, headless, gui, output_filename):
    prevent_sleep_windows()
    driver = None
    try:
        gui.log("Starting browser...")
        driver = setup_driver(headless=headless)
        gui.log("Browser started. Navigating to Google Maps...")

        listings = collect_all_listings(driver, gui, base_query)
        try:
            driver.quit()
        except:
            pass
        driver = None

        gui.log(f"Collected {len(listings)} entries. Building DataFrame...")
        df = pd.DataFrame(listings)

        emails_list = []
        for i, website in enumerate(df.get("Website", [])):
            gui.update_status(f"Scraping emails: {i+1}/{len(df)}")
            em = find_emails_from_website(website)
            emails_list.append(em)
            gui.log(f"Website: {website} → Email: {em}")
            time.sleep(0.5)

        df["ScrapedEmail"] = emails_list
        df.to_excel(output_filename, index=False)
        gui.log(f"Saved results to {output_filename}")
        gui.done(output_filename)
    except Exception as e:
        gui.error(str(e))
    finally:
        clear_sleep_windows()
        if driver:
            try:
                driver.quit()
            except:
                pass

# ---------- GUI class ----------
class VisionMateGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("VisionMate Automation Tool")

        self.output_filename = os.path.join(os.getcwd(), "visionmate_results.xlsx")

        # Query + options
        frame_top = tk.Frame(root)
        frame_top.pack(fill="x", padx=10, pady=10)

        tk.Label(frame_top, text="Search Query:").pack(side="left")
        self.query_var = tk.StringVar()
        tk.Entry(frame_top, textvariable=self.query_var, width=40).pack(side="left", padx=5)

        self.headless_var = tk.BooleanVar()
        tk.Checkbutton(frame_top, text="Run headless", variable=self.headless_var).pack(side="left")

        tk.Button(frame_top, text="Start", command=self.start_scrape).pack(side="left", padx=5)
        tk.Button(frame_top, text="Quit", command=self.quit).pack(side="left")

        # Log box
        self.log_box = tk.Text(root, height=15, width=80, state="disabled")
        self.log_box.pack(padx=10, pady=5)

        # Status + download
        frame_bottom = tk.Frame(root)
        frame_bottom.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_bottom, text="Status:").pack(side="left")
        self.status_var = tk.StringVar()
        tk.Label(frame_bottom, textvariable=self.status_var).pack(side="left", padx=5)
        self.download_btn = tk.Button(frame_bottom, text="Download Results", command=self.download, state="disabled")
        self.download_btn.pack(side="right")

        self.overlay = None

    def log(self, msg):
        self.log_box.config(state="normal")
        self.log_box.insert("end", time.strftime('%H:%M:%S') + "  " + msg + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    def update_status(self, msg):
        self.status_var.set(msg)
        if self.overlay:
            self.overlay_label.config(text=msg)
            self.overlay.update()

    def start_scrape(self):
        query = self.query_var.get().strip()
        if not query:
            messagebox.showwarning("Warning", "Please enter a search query.")
            return
        if not messagebox.askokcancel("Confirm", "Please set system sleep to at least 30 mins.\nDo not touch machine while running.\n\nProceed?"):
            return
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")
        self.status_var.set("Running...")

        self.show_overlay("Processing... 🔒 Do not touch the screen.")

        t = threading.Thread(target=worker_scrape, args=(query, self.headless_var.get(), self, self.output_filename), daemon=True)
        t.start()

    def show_overlay(self, msg):
        self.overlay = tk.Toplevel(self.root)
        self.overlay.title("")
        self.overlay.geometry("300x120+200+200")
        self.overlay.attributes("-alpha", 0.9)
        self.overlay.attributes("-topmost", True)
        self.overlay.grab_set()

        tk.Label(self.overlay, text="🔒", font=("Helvetica", 30)).pack(pady=5)
        self.overlay_label = tk.Label(self.overlay, text=msg)
        self.overlay_label.pack(pady=5)
        ttk.Progressbar(self.overlay, mode="indeterminate").pack(fill="x", padx=20, pady=10)
        self.overlay.after(100, lambda: self.overlay.children['!progressbar'].start())

    def close_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None

    def done(self, result_file):
        self.close_overlay()
        self.status_var.set(f"Completed. File: {os.path.basename(result_file)}")
        self.download_btn.config(state="normal")
        messagebox.showinfo("Done", "Scrape completed successfully! Click Download Results.")

    def error(self, msg):
        self.close_overlay()
        self.status_var.set("Error")
        messagebox.showerror("Error", msg)

    def download(self):
        folder = filedialog.askdirectory(title="Choose folder to save results")
        if folder:
            try:
                dest_path = os.path.join(folder, os.path.basename(self.output_filename))
                if os.path.exists(dest_path):
                    if not messagebox.askyesno("Overwrite?", f"{dest_path} exists. Overwrite?"):
                        return
                shutil.move(self.output_filename, dest_path)
                messagebox.showinfo("Saved", f"File moved to: {dest_path}")
                self.download_btn.config(state="disabled")
                self.status_var.set("Idle")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def quit(self):
        if self.overlay:
            messagebox.showwarning("Wait", "Automation is running. Please wait until it completes.")
            return
        self.root.quit()

# ---------- main ----------
if __name__ == "__main__":
    root = tk.Tk()
    gui = VisionMateGUI(root)
    root.mainloop()