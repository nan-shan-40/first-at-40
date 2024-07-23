#this code pulls repo names, language composition from gitHub
from github import Github
import openpyxl
from openpyxl.utils import get_column_letter

# Replace with your GitHub personal access token
GITHUB_TOKEN = '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'
ORG_NAME = 'octokit'

def get_repo_info():
    g = Github(GITHUB_TOKEN)
    org = g.get_organization(ORG_NAME)
    repos = org.get_repos()

    data = []
    for repo in repos:
        repo_name = repo.name
        last_updated = repo.updated_at.strftime('%Y-%m-%d %H:%M:%S')
        languages = repo.get_languages()

        total_bytes = sum(languages.values())
        if total_bytes > 0:
            language_composition = {k: (v / total_bytes) * 100 for k, v in languages.items()}
            primary_language = max(language_composition, key=language_composition.get)
        else:
            language_composition = {}
            primary_language = 'N/A'

        language_composition_str = ', '.join([f'{lang}: {bytes} bytes' for lang, bytes in languages.items()])
        data.append([repo_name, last_updated, language_composition_str, primary_language])

    return data

def write_to_excel(data):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Repositories"

    headers = ["Repository Name", "Last Updated", "Languages", "Primary Language"]
    sheet.append(headers)

    for row in data:
        sheet.append(row)

    # Auto adjust column widths
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column].width = adjusted_width

    workbook.save("GitHubRepositories.xlsx")

if __name__ == "__main__":
    repo_data = get_repo_info()
    write_to_excel(repo_data)
