import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
from docx import Document
from openai import OpenAI

client = OpenAI()

class ContentDownloader:
    def __init__(self, service_account_file, spreadsheet_id, sheet_name):
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(service_account_file, scope)
        self.client = gspread.authorize(creds)
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name

    def fetch_data(self):
        sheet = self.client.open_by_key(self.spreadsheet_id).worksheet(self.sheet_name)
        return sheet.get_all_records()

    def download_html(self, url):
        try:
            r = requests.get(url, timeout=10, headers={'User-Agent': 'Mozilla/5.0'})
            r.raise_for_status()
            return r.text
        except:
            return ""

    def get_all_pages(self, data):
        results = {}
        for row in data:
            site = row['url site']
            urls = eval(row['URL serp (по убыванию частоты)']) if isinstance(row['URL serp (по убыванию частоты)'], str) else row['URL serp (по убыванию частоты)']
            html_list = []
            for url in urls:
                html = self.download_html(url)
                if html:
                    html_list.append(html)
            results[site] = html_list
        return results

class ChatGPTAnalyzer:
    def analyze(self, pages_dict, keywords_data):
        analysis_results = {}
        for site, html_list in pages_dict.items():
            content_chunks = html_list[:3]
            content_sample = "\n".join(content_chunks)
            prompt = (
                "Проанализируй HTML-тексты и посчитай:\n"
                "- среднее количество вхождений ключевых фраз во всех словоформах\n"
                "- среднее количество вхождений отдельных слов из ключевых фраз\n"
                "- средний объем текста\n"
                "- список 20 LSI слов\n"
                f"\nКлючевые фразы: {', '.join([row['Фраза'] for row in keywords_data if row['url site']==site])}\n"
                f"\nТекст:\n{content_sample[:12000]}"
            )
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "Ты помощник для SEO анализа."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0
            )
            text = response.choices[0].message.content
            analysis_results[site] = {
                'mean_occurrences': 0,
                'mean_word_occurrences': 0,
                'mean_length': 0,
                'lsi': []
            }
            lines = text.split('\n')
            for line in lines:
                if "вхождений ключевых фраз" in line:
                    analysis_results[site]['mean_occurrences'] = int(''.join(filter(str.isdigit, line)))
                elif "вхождений отдельных слов" in line:
                    analysis_results[site]['mean_word_occurrences'] = int(''.join(filter(str.isdigit, line)))
                elif "средний объем" in line:
                    analysis_results[site]['mean_length'] = int(''.join(filter(str.isdigit, line)))
                elif "LSI" in line:
                    lsi = line.split(':')[-1].strip()
                    analysis_results[site]['lsi'] = [w.strip() for w in lsi.split(',')]
        return analysis_results

class ReportGenerator:
    def generate_report(self, analysis_results):
        doc = Document()
        doc.add_heading('Content Optimization Recommendations', 0)
        for site, stats in analysis_results.items():
            doc.add_heading(site, level=1)
            doc.add_paragraph(f"Среднее количество вхождений запросов: {stats['mean_occurrences']}")
            doc.add_paragraph(f"Среднее количество вхождений отдельных слов: {stats['mean_word_occurrences']}")
            doc.add_paragraph(f"Средний объем текста: {stats['mean_length']}")
            doc.add_paragraph("LSI слова: " + ', '.join(stats['lsi']))
        doc.save('recommendations.docx')

if __name__ == '__main__':
    downloader = ContentDownloader('service_account.json', 'SPREADSHEET_ID', 'CheckTop')
    data = downloader.fetch_data()
    pages = downloader.get_all_pages(data)

    analyzer = ChatGPTAnalyzer()
    analysis_results = analyzer.analyze(pages, data)

    report = ReportGenerator()
    report.generate_report(analysis_results)
