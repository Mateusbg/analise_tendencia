import csv
import json
import os
import sys
import webbrowser
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import tkinter as tk
from tkinter import filedialog

# Carregar configurações
config_path = os.path.join(os.path.dirname(__file__), 'config.json')
try:
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
except FileNotFoundError:
    print("Aviso: config.json não encontrado. Usando configurações padrão.")
    config = {
        "encoding": "utf-8",
        "separator": ",",
        "auto_detect_separator": True,
        "display": {"max_rows": None, "max_columns": None, "max_colwidth": None},
        "filtro_ativo": True
    }


def detect_separator(file_path, encoding='utf-8'):
    try:
        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
            sample = f.read(8192)
            sniffer = csv.Sniffer()
            dialect = sniffer.sniff(sample, delimiters=[',', ';', '\t', '|'])
            return dialect.delimiter
    except Exception:
        return None


def read_csv_file(file_path, encoding='utf-8', default_sep=','):
    sep = default_sep
    if config.get('auto_detect_separator', True):
        detected = detect_separator(file_path, encoding=encoding)
        if detected:
            sep = detected
            print(f"Separador detectado: '{sep}'")
        else:
            print("Não foi possível detectar o separador automaticamente. Usando separador padrão ','")
    try:
        return pd.read_csv(file_path, sep=sep, encoding=encoding, engine='python')
    except UnicodeDecodeError:
        return pd.read_csv(file_path, sep=sep, encoding='latin1', engine='python')
    except pd.errors.ParserError:
        for alt_sep in [',', ';', '\t', '|']:
            if alt_sep == sep:
                continue
            try:
                print(f"Tentando leitura com separador '{alt_sep}'")
                return pd.read_csv(file_path, sep=alt_sep, encoding=encoding, engine='python')
            except Exception:
                continue
        raise


def find_date_column(df):
    candidates = [col for col in df.columns if isinstance(col, str) and any(keyword in col.lower() for keyword in ('data', 'date', 'venc', 'due', 'prazo', 'devolu'))]
    for col in candidates:
        parsed = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
        if parsed.notna().any():
            return col
    return candidates[0] if candidates else None


def create_table_trace(df, name):
    header = [str(col) for col in df.columns]
    cells = [df[col].astype(str).tolist() for col in df.columns]
    return go.Table(
        header=dict(values=header, fill_color='paleturquoise', align='left'),
        cells=dict(values=cells, fill_color='lavender', align='left'),
        name=name
    )


def create_plotly_visualization(df, output_path):
    df_full = df.copy()
    max_rows = 500
    if len(df_full) > max_rows:
        print('A visualização foi limitada a 500 linhas para manter a página leve.')

    date_col = find_date_column(df_full)
    overdue_trace = None
    traces = []
    buttons = []
    export_data = {
        'all': df_full.to_dict('records'),
        'overdue': []
    }

    if date_col is not None:
        date_series = pd.to_datetime(df_full[date_col], errors='coerce', dayfirst=True)
        today = pd.Timestamp.now().normalize()
        overdue_mask = date_series + pd.Timedelta(days=10) < today
        df_overdue = df_full[overdue_mask]

        if len(df_overdue) > 0:
            df_overdue_display = df_overdue.head(max_rows)
            overdue_trace = create_table_trace(df_overdue_display, 'Atrasados >10 dias')
            traces.append(overdue_trace)
            export_data['overdue'] = df_overdue.to_dict('records')
            buttons.append(dict(
                label='Atrasados >10 dias',
                method='update',
                args=[{'visible': [False, True]}, {'title': 'Linhas com atraso maior que 10 dias'}]
            ))
    else:
        pass

    df_display = df_full.head(max_rows)
    main_trace = create_table_trace(df_display, 'Todos')
    traces.insert(0, main_trace)
    buttons.insert(0, dict(
        label='Todos',
        method='update',
        args=[{'visible': [True] + ([False] if overdue_trace is not None else [])}, {'title': 'Todas as linhas'}]
    ))

    fig = go.Figure(data=traces)
    if buttons:
        fig.update_layout(
            updatemenus=[dict(
                type='buttons',
                direction='right',
                active=0,
                x=0.5,
                xanchor='center',
                y=1.15,
                buttons=buttons
            )]
        )
    fig.update_layout(
        title_text='Visualização do Arquivo CSV',
        title_x=0.5,
        margin=dict(l=10, r=10, t=120, b=10),
        autosize=True
    )

    plot_html = pio.to_html(fig, full_html=False, include_plotlyjs='cdn')
    export_json = json.dumps(export_data, default=str)
    disabled_attr = ' disabled' if not export_data['overdue'] else ''
    export_button_html = f"""
    <div style='text-align:center; margin-bottom:15px;'>
      <button onclick=\"exportToExcel('all')\" style='margin-right:10px; padding:10px 18px; font-size:14px;'>Exportar todos para Excel</button>
      <button onclick=\"exportToExcel('overdue')\" style='padding:10px 18px; font-size:14px;' {disabled_attr}>Exportar atrasados > 10 dias para Excel</button>
    </div>
    """
    custom_html = """
<!DOCTYPE html>
<html>
<head>
  <meta charset='utf-8' />
  <title>Visualização CSV</title>
  <script src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'></script>
</head>
<body>
  <h2 style='text-align:center;'>Visualização do Arquivo CSV</h2>
  @@EXPORT_BUTTON_HTML@@
  @@PLOT_HTML@@
  <script>
    const exportData = @@EXPORT_JSON@@;
    function exportToExcel(view) {
      const data = exportData[view] || [];
      if (!data.length) {
        alert('Nenhum registro disponível para exportação.');
        return;
      }
      const worksheet = XLSX.utils.json_to_sheet(data);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, view === 'all' ? 'Todos' : 'Atrasados');
      const filename = view === 'all' ? 'visualizacao_todos.xlsx' : 'visualizacao_atrasados.xlsx';
      XLSX.writeFile(workbook, filename);
    }
  </script>
</body>
</html>
"""
    custom_html = custom_html.replace('@@EXPORT_BUTTON_HTML@@', export_button_html)
    custom_html = custom_html.replace('@@PLOT_HTML@@', plot_html)
    custom_html = custom_html.replace('@@EXPORT_JSON@@', export_json)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(custom_html)
    file_url = 'file://' + os.path.abspath(output_path).replace('\\', '/')
    webbrowser.open_new_tab(file_url)
    print(f"Visualização aberta no navegador: {output_path}")


root = tk.Tk()
root.withdraw()
root.update()
root.call('wm', 'attributes', '.', '-topmost', True)

caminho_arquivo = filedialog.askopenfilename(
    parent=root,
    title="Selecione o arquivo CSV",
    filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
)

root.destroy()

if not caminho_arquivo:
    print("Nenhum arquivo foi selecionado.")
    sys.exit()

print(f"Arquivo selecionado: {caminho_arquivo}")
encoding = config.get("encoding", "utf-8")
separator = config.get("separator", ",")

try:
    df = read_csv_file(caminho_arquivo, encoding=encoding, default_sep=separator)
except Exception as error:
    print(f"Erro ao ler o arquivo CSV: {error}")
    sys.exit(1)

if df.columns.dtype == object:
    df.columns = df.columns.str.strip()

# Aplicar configurações de display
display_config = config.get("display", {})
pd.set_option("display.max_rows", display_config.get("max_rows"))
pd.set_option("display.max_columns", display_config.get("max_columns"))
pd.set_option("display.width", None)
pd.set_option("display.max_colwidth", display_config.get("max_colwidth"))
pd.set_option("display.expand_frame_repr", False)

output_html = os.path.join(os.path.dirname(__file__), 'csv_visualizacao.html')
create_plotly_visualization(df, output_html)
