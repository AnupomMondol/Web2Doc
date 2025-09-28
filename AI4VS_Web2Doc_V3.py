#!/usr/bin/env python3
import os
import re
import time
import base64
import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from urllib.parse import urlparse, parse_qs, unquote, quote

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager
from PyPDF2 import PdfMerger


class ConcordConverter:
    TMP_DIR = "tmp_pdfs"
    PAPER_WIDTH = 12.5            # inches
    MARGINS = dict(
        marginTop=0.5,
        marginBottom=0.5,
        marginLeft=0.5,
        marginRight=0.5
    )
    MAX_WAIT = 30                 # seconds for page load
    PX_PER_INCH = 96.0

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Concord Platform Converter")
        self.root.geometry("800x600")
        self._build_gui()

    # ──────────────────────────────────────────────────────────────────────
    # GUI
    # ──────────────────────────────────────────────────────────────────────
    def _build_gui(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        # URL entry
        ttk.Label(frm, text="Concord Platform URL:")\
            .grid(row=0, column=0, sticky=tk.W, pady=5)
        self.url_entry = ttk.Entry(frm, width=70)
        self.url_entry.grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=5)

        # Output folder
        ttk.Label(frm, text="Output Folder:")\
            .grid(row=1, column=0, sticky=tk.W, pady=5)
        self.folder_path = tk.StringVar()
        ttk.Entry(frm, textvariable=self.folder_path, width=50)\
            .grid(row=1, column=1, sticky=tk.EW, pady=5)
        ttk.Button(frm, text="Browse", command=self._browse_folder)\
            .grid(row=1, column=2, padx=5)

        # Conversion options
        ofrm = ttk.LabelFrame(frm, text="Conversion Options", padding=10)
        ofrm.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=10)
        self.convert_pdf_var = tk.BooleanVar(value=True)
        self.convert_docx_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(ofrm, text="Convert to PDF", variable=self.convert_pdf_var)\
            .pack(anchor=tk.W)
        ttk.Checkbutton(ofrm, text="Convert to DOCX", variable=self.convert_docx_var)\
            .pack(anchor=tk.W)

        # Convert button
        ttk.Button(frm, text="Convert", command=self.start_conversion)\
            .grid(row=3, column=0, columnspan=3, pady=10)

        # Log console
        clog = ttk.LabelFrame(frm, text="Conversion Log", padding=10)
        clog.grid(row=4, column=0, columnspan=3, sticky=tk.NSEW)
        self.console = scrolledtext.ScrolledText(
            clog, width=90, height=15, state='disabled'
        )
        self.console.pack(fill=tk.BOTH, expand=True)

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(4, weight=1)

    def _browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)

    def log(self, msg: str):
        self.console.config(state='normal')
        self.console.insert(tk.END, msg + "\n")
        self.console.config(state='disabled')
        self.console.see(tk.END)
        self.root.update()

    # ──────────────────────────────────────────────────────────────────────
    # Helpers
    # ──────────────────────────────────────────────────────────────────────
    @staticmethod
    def sanitize_filename(name: str) -> str:
        return re.sub(r'[<>:"/\\|?*\x00-\x1F]', '_', name)

    def ensure_dirs(self, *paths: str):
        for p in paths:
            os.makedirs(p, exist_ok=True)

    def get_driver(self) -> webdriver.Chrome:
        opts = Options()
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--disable-extensions")
        opts.add_argument("--disable-software-rasterizer")
        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opts
        )

    def extract_activity_info(self, url: str) -> dict:
        parsed = urlparse(url)
        qs = parse_qs(parsed.query, keep_blank_values=True)
        param_order = [p.split('=')[0] for p in parsed.query.split('&') if '=' in p]

        if 'activity' in qs:
            act_url = unquote(qs['activity'][0])
            aid = re.search(r'activities/(\d+)\.json', act_url).group(1)
            return dict(
                format='activity',
                activity_url=act_url,
                activity_id=aid,
                run_key=qs.get('runKey', [None])[0],
                page=qs.get('page', [None])[0],
                param_order=param_order
            )

        seq_url = unquote(qs.get('sequence', [''])[0])
        sid = re.search(r'sequences/(\d+)\.json', seq_url).group(1)
        return dict(
            format='sequence',
            sequence_url=seq_url,
            sequence_id=sid,
            activity_id=qs.get('sequenceActivity', [''])[0],
            preview='preview' in qs,
            page=qs.get('page', [None])[0],
            param_order=param_order
        )

    def build_page_url(self, url_info, page_id=None):
        """Build page URL preserving the exact parameter order from input URL"""
        if url_info['format'] == 'activity':
            # Activity format: activity -> page -> runKey
            params = {
                'activity': quote(url_info['activity_url'], safe=''),
                'page': page_id,
                'runKey': url_info['run_key']
            }
            # Remove None values
            params = {k: v for k, v in params.items() if v is not None}
            return "https://activity-player.concord.org/?" + "&".join(f"{k}={v}" for k, v in params.items())
        else:
            # Sequence format: page -> preview -> sequence -> sequenceActivity
            params = {
                'page': page_id,
                'preview': '',
                'sequence': quote(url_info['sequence_url'], safe=''),
                'sequenceActivity': url_info['activity_id']
            }
            # Remove None values
            params = {k: v for k, v in params.items() if v is not None}
            return "https://activity-player.concord.org/?" + "&".join(
                k if v == '' else f"{k}={v}" for k, v in params.items()
            )

    def fetch_metadata(self, info: dict) -> tuple[str, list]:
        if info['format'] == 'sequence':
            data = requests.get(info['sequence_url']).json()
            act = next(
                (a for a in data.get('activities', [])
                 if str(a['id']) == info['activity_id'].lstrip('activity_')),
                None
            )
            if not act:
                raise ValueError("Activity not found in sequence")
            return act.get('name', f"Activity_{info['activity_id']}"), act.get('pages', [])
        else:
            data = requests.get(info['activity_url']).json()
            return data.get('name', "Activity"), data.get('pages', [])

    def save_page_pdf(self, driver, url: str, out_path: str) -> bool:
        """
        Renders the activity-player page at `url` to PDF, sizing height
        dynamically to the content's full scrollHeight.
        """
        try:
            self.log(f" → Rendering: {url}")
            driver.get(url)

            # wait for the real scroll container
            container = WebDriverWait(driver, self.MAX_WAIT).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div.app[data-cy='app']"))
            )
            time.sleep(1)  # let dynamic content finish

            # scroll to bottom repeatedly
            last_h = 0
            for _ in range(20):
                last_h = driver.execute_script(
                    "arguments[0].scrollTo(0, arguments[0].scrollHeight);"
                    "return arguments[0].scrollHeight;",
                    container
                )
                time.sleep(0.3)

            # convert px→inches for PDF
            height_in = (last_h / self.PX_PER_INCH
                         + self.MARGINS["marginTop"]
                         + self.MARGINS["marginBottom"])

            # print to PDF with dynamic height
            pdf = driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "paperWidth": self.PAPER_WIDTH,
                "paperHeight": height_in,
                **self.MARGINS,
                "displayHeaderFooter": True,
                "headerTemplate": (
                    "<div style='font-size:10px;text-align:center;'>Activity</div>"
                ),
                "footerTemplate": (
                    "<div style='font-size:10px;text-align:center;'>"
                    "<span class='pageNumber'></span> / "
                    "<span class='totalPages'></span></div>"
                )
            })

            with open(out_path, "wb") as f:
                f.write(base64.b64decode(pdf['data']))
            return True

        except Exception as e:
            self.log(f"[ERROR] PDF failed: {e}")
            return False

    @staticmethod
    def get_unique_filepath(path: str) -> str:
        """
        If `path` exists, returns path with a numeric suffix:
        e.g. /foo/bar.pdf → /foo/bar(1).pdf, /foo/bar(2).pdf, …
        """
        base, ext = os.path.splitext(path)
        counter = 1
        candidate = path
        while os.path.exists(candidate):
            candidate = f"{base}({counter}){ext}"
            counter += 1
        return candidate

    def render_and_merge(
            self,
            driver,
            info: dict,
            pages: list,
            safe_title: str,
            out_dir: str
    ) -> str:
        """
        Renders all pages to individual PDFs (tmp_pdfs/...), then merges them
        in the order (home first, then page1, page2, …) into safe_title.pdf
        in out_dir. Handles PermissionErrors on write/delete gracefully.
        """
        # 0) ensure both temp and output dirs exist
        self.ensure_dirs(self.TMP_DIR, out_dir)

        temp_files = []

        # 1) home page
        home_pdf = os.path.join(self.TMP_DIR, f"{safe_title}_home.pdf")
        if self.save_page_pdf(driver, self.build_page_url(info), home_pdf):
            temp_files.append(home_pdf)

        # 2) content pages
        for idx, pg in enumerate(pages, start=1):
            pid = pg.get('id', str(idx))
            pdf_path = os.path.join(self.TMP_DIR, f"{safe_title}_page{pid}.pdf")
            url = self.build_page_url(info, f"page_{pid}")
            if self.save_page_pdf(driver, url, pdf_path):
                temp_files.append(pdf_path)

        if not temp_files:
            raise RuntimeError("No pages rendered; aborting merge.")

        # 3) sort: home first, then numeric pages
        def _sort_key(p: str):
            n = os.path.basename(p).lower()
            if n.endswith("_home.pdf"):
                return (0, 0)
            m = re.search(r'_page(\d+)\.pdf$', n)
            if m:
                return (1, int(m.group(1)))
            return (2, n)

        temp_files.sort(key=_sort_key)

        # 4) prepare final path, but don’t overwrite an in‑use file
        raw_final = os.path.join(out_dir, f"{safe_title}.pdf")
        final_pdf = self.get_unique_filepath(raw_final)

        # 5) merge into final_pdf …
        merger = PdfMerger()
        for fpath in temp_files:
            merger.append(fpath)

        try:
            with open(final_pdf, "wb") as out_f:
                merger.write(out_f)
        except Exception as e:
            self.log(f"[ERROR] Failed to write '{final_pdf}': {e}")
            raise
        finally:
            merger.close()

        # 6) cleanup temp files
        for fpath in temp_files:
            try:
                os.remove(fpath)
            except PermissionError as pe:
                self.log(f"[WARN] Permission denied deleting '{fpath}': {pe}")
            except Exception:
                pass

        return final_pdf

    # def pdf_to_docx(self, pdf_path: str, out_dir: str) -> str | None:
    #     self.log("→ Starting DOCX conversion…")
    #     driver = self.get_driver()
    #     try:
    #         driver.get("https://www.ilovepdf.com/pdf_to_word")
    #         try:
    #             WebDriverWait(driver, 10).until(
    #                 EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Accept')]"))
    #             ).click()
    #         except: pass
    #
    #         upload = WebDriverWait(driver, 20).until(
    #             EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
    #         )
    #         upload.send_keys(os.path.abspath(pdf_path))
    #
    #         WebDriverWait(driver, 60).until(
    #             EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Convert to WORD')]"))
    #         ).click()
    #
    #         dl = WebDriverWait(driver, 120).until(
    #             EC.element_to_be_clickable((By.XPATH, "//a[contains(.,'Download WORD')]"))
    #         )
    #         url = dl.get_attribute("href")
    #         data = requests.get(url).content
    #
    #         docx = os.path.join(
    #             out_dir,
    #             os.path.splitext(os.path.basename(pdf_path))[0] + ".docx"
    #         )
    #         with open(docx, "wb") as f:
    #             f.write(data)
    #         return docx
    #
    #     except Exception as e:
    #         self.log(f"[ERROR] DOCX failed: {e}")
    #         return None
    #
    #     finally:
    #         driver.quit()


    def pdf_to_docx(self, pdf_path: str, out_dir: str) -> str | None:
        self.log("→ Starting DOCX conversion…")
        driver = self.get_driver()
        try:
            driver.get("https://www.ilovepdf.com/pdf_to_word")
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Accept')]"))
            ).click()
        except: pass

        upload = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
        )
        upload.send_keys(os.path.abspath(pdf_path))

        WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Convert to WORD')]"))
        ).click()

        dl = WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(.,'Download WORD')]"))
        )
        url = dl.get_attribute("href")
        data = requests.get(url).content
        docx = os.path.join(out_dir, os.path.splitext(os.path.basename(pdf_path))[0] + ".docx")
        with open(docx, "wb") as f:
            f.write(data)
        driver.quit()
        return docx

    # ──────────────────────────────────────────────────────────────────────
    # Main entry
    # ──────────────────────────────────────────────────────────────────────
    def start_conversion(self):
        url = self.url_entry.get().strip()
        out_dir = self.folder_path.get().strip() or os.getcwd()

        if not url.lower().startswith(("http://", "https://")):
            return messagebox.showerror(
                "Error", "Invalid URL; must start with http:// or https://"
            )

        self.ensure_dirs(out_dir, self.TMP_DIR)
        info = self.extract_activity_info(url)
        driver = self.get_driver()

        try:
            title, pages = self.fetch_metadata(info)
            safe_title = self.sanitize_filename(title)
            self.log(f"[INFO] Activity: {title} ({len(pages)} pages)")

            pdf_path = self.render_and_merge(
                driver, info, pages, safe_title, out_dir
            )
            self.log(f"[OK] PDF saved: {pdf_path}")

            if self.convert_docx_var.get():
                docx_path = self.pdf_to_docx(pdf_path, out_dir)
                if docx_path:
                    self.log(f"[OK] DOCX saved: {docx_path}")
                    messagebox.showinfo(
                        "Success",
                        f"Saved:\nPDF: {pdf_path}\nDOCX: {docx_path}"
                    )
                else:
                    raise RuntimeError("DOCX conversion failed")
            else:
                messagebox.showinfo("Success", f"PDF saved to:\n{pdf_path}")

        except Exception as e:
            self.log(f"[ERROR] {e}")
            messagebox.showerror("Conversion failed", str(e))
        finally:
            driver.quit()


if __name__ == "__main__":
    root = tk.Tk()
    ConcordConverter(root)
    root.mainloop()
