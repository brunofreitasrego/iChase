from pypdf import PdfReader
import os
import pandas as pd
import re
from html.parser import HTMLParser
import subprocess
import os
import codecs
import win32com.client
from docx import Document

def list_files_in_current_directory():
    """List all files in the current directory."""
    return [f for f in os.listdir(".") if os.path.isfile(os.path.join(".", f))]

def is_file_empty(file_path):
    return os.stat(file_path).st_size == 0

has_cif_file = False
jar_file = ""

for file_name in list_files_in_current_directory():
    _, ext = os.path.splitext(file_name)
    if ext.lower() == '.jar':
        jar_file = file_name
    if ext.lower() == '.cif':
        has_cif_file = True

if has_cif_file:
    if jar_file:
        command = [
            'java',
            '-jar',
            './' + jar_file,
            './'
        ]
        subprocess.run(command, capture_output=True, text=True)
        file_names = list_files_in_current_directory()
        failed_cif_num = 0
        for file_name in file_names:
            file_name_no_ext, ext = os.path.splitext(file_name)
            if ext.lower() == '.cif' and (file_name_no_ext not in file_names or is_file_empty(file_name_no_ext)):
                failed_cif_num = failed_cif_num + 1
                if failed_cif_num == 1:
                    print("Segue a lista de arquivos CIF que não tiveram o CIF removido corretamente! (tente remover o primeiro arquivo da lista e rodar novamente)")
                print(file_name)
        

    else:
        print("Existem arquivos CIF mas está faltando o programa Decrypt!")

destinations = frozenset((
    'aftncn','aftnsep','aftnsepc','annotation','atnauthor','atndate','atnicn','atnid',
    'atnparent','atnref','atntime','atrfend','atrfstart','author','background',
    'bkmkend','bkmkstart','blipuid','buptim','category','colorschememapping',
    'colortbl','comment','company','creatim','datafield','datastore','defchp','defpap',
    'do','doccomm','docvar','dptxbxtext','ebcend','ebcstart','factoidname','falt',
    'fchars','ffdeftext','ffentrymcr','ffexitmcr','ffformat','ffhelptext','ffl',
    'ffname','ffstattext','file','filetbl','fldinst','fldtype',
    'fname','fontemb','fontfile','fonttbl','footer','footerf','footerl','footerr',
    'footnote','formfield','ftncn','ftnsep','ftnsepc','g','generator','gridtbl',
    'header','headerf','headerl','headerr','hl','hlfr','hlinkbase','hlloc','hlsrc',
    'hsv','htmltag','info','keycode','keywords','latentstyles','lchars','levelnumbers',
    'leveltext','lfolevel','linkval','list','listlevel','listname','listoverride',
    'listoverridetable','listpicture','liststylename','listtable','listtext',
    'lsdlockedexcept','macc','maccPr','mailmerge','maln','malnScr','manager','margPr',
    'mbar','mbarPr','mbaseJc','mbegChr','mborderBox','mborderBoxPr','mbox','mboxPr',
    'mchr','mcount','mctrlPr','md','mdeg','mdegHide','mden','mdiff','mdPr','me',
    'mendChr','meqArr','meqArrPr','mf','mfName','mfPr','mfunc','mfuncPr','mgroupChr',
    'mgroupChrPr','mgrow','mhideBot','mhideLeft','mhideRight','mhideTop','mhtmltag',
    'mlim','mlimloc','mlimlow','mlimlowPr','mlimupp','mlimuppPr','mm','mmaddfieldname',
    'mmath','mmathPict','mmathPr','mmaxdist','mmc','mmcJc','mmconnectstr',
    'mmconnectstrdata','mmcPr','mmcs','mmdatasource','mmheadersource','mmmailsubject',
    'mmodso','mmodsofilter','mmodsofldmpdata','mmodsomappedname','mmodsoname',
    'mmodsorecipdata','mmodsosort','mmodsosrc','mmodsotable','mmodsoudl',
    'mmodsoudldata','mmodsouniquetag','mmPr','mmquery','mmr','mnary','mnaryPr',
    'mnoBreak','mnum','mobjDist','moMath','moMathPara','moMathParaPr','mopEmu',
    'mphant','mphantPr','mplcHide','mpos','mr','mrad','mradPr','mrPr','msepChr',
    'mshow','mshp','msPre','msPrePr','msSub','msSubPr','msSubSup','msSubSupPr','msSup',
    'msSupPr','mstrikeBLTR','mstrikeH','mstrikeTLBR','mstrikeV','msub','msubHide',
    'msup','msupHide','mtransp','mtype','mvertJc','mvfmf','mvfml','mvtof','mvtol',
    'mzeroAsc','mzeroDesc','mzeroWid','nesttableprops','nextfile','nonesttables',
    'objalias','objclass','objdata','object','objname','objsect','objtime','oldcprops',
    'oldpprops','oldsprops','oldtprops','oleclsid','operator','panose','password',
    'passwordhash','pgp','pgptbl','picprop','pict','pn','pnseclvl','pntext','pntxta',
    'pntxtb','printim','private','propname','protend','protstart','protusertbl','pxe',
    'result','revtbl','revtim','rsidtbl','rxe','shp','shpgrp','shpinst',
    'shppict','shprslt','shptxt','sn','sp','staticval','stylesheet','subject','sv',
    'svb','tc','template','themedata','title','txe','ud','upr','userprops',
    'wgrffmtfilter','windowcaption','writereservation','writereservhash','xe','xform',
    'xmlattrname','xmlattrvalue','xmlclose','xmlname','xmlnstbl',
    'xmlopen',
    ))
# fmt: on


# Translation of some special characters.
specialchars = {
    "par": "\n",
    "sect": "\n\n",
    "page": "\n\n",
    "line": "\n",
    "tab": "\t",
    "emdash": "\u2014",
    "endash": "\u2013",
    "emspace": "\u2003",
    "enspace": "\u2002",
    "qmspace": "\u2005",
    "bullet": "\u2022",
    "lquote": "\u2018",
    "rquote": "\u2019",
    "ldblquote": "\u201C",
    "rdblquote": "\u201D",
    "row": "\n",
    "cell": "|",
    "nestcell": "|",
    "~": "\xa0",
    "\n":"\n",
    "\r": "\r",
    "{": "{",
    "}": "}",
    "\\": "\\",
    "-": "\xad",
    "_": "\u2011"

}

PATTERN = re.compile(
    r"\\([a-z]{1,32})(-?\d{1,10})?[ ]?|\\'([0-9a-f]{2})|\\([^a-z])|([{}])|[\r\n]+|(.)",
    re.IGNORECASE,
)

HYPERLINKS = re.compile(
    r"(\{\\field\{\s*\\\*\\fldinst\{.*HYPERLINK\s(\".*\")\}{2}\s*\{.*?\s+(.*?)\}{2,3})",
    re.IGNORECASE
)

    
def format_rtf_line(text, encoding="cp1252", errors="strict"):
    """ Converts the rtf text to plain text.

    Parameters
    ----------
    text : str
        The rtf text
    encoding : str
        Input encoding which is ignored if the rtf file contains an explicit codepage directive, 
        as it is typically the case. Defaults to `cp1252` encoding as it the most commonly used.
    errors : str
        How to handle encoding errors. Default is "strict", which throws an error. Another
        option is "ignore" which, as the name says, ignores encoding errors.

    Returns
    -------
    str
        the converted rtf text as a python unicode string
    """
    text = re.sub(HYPERLINKS, "\\1(\\2)", text) # captures links like link_text(http://link_dest)
    stack = []
    ignorable = False  # Whether this group (and all inside it) are "ignorable".
    ucskip = 1  # Number of ASCII characters to skip after a unicode character.
    curskip = 0  # Number of ASCII characters left to skip
    hexes = None
    out = ''

    for match in PATTERN.finditer(text):
        word, arg, _hex, char, brace, tchar = match.groups()
        if hexes and not _hex:
            try:
                out += bytes.fromhex(hexes).decode(encoding=encoding, errors=errors)
            except UnicodeDecodeError:
                out += bytes.fromhex(hexes).decode('latin1', errors='ignore')
            hexes = None
        if brace:
            curskip = 0
            if brace == "{":
                # Push state
                stack.append((ucskip, ignorable))
            elif brace == "}":
                # Pop state
                if stack:
                    ucskip, ignorable = stack.pop()
                # sample_3.rtf throws an IndexError because of stack being empty.
                # don't know right now how this could happen, so for now this is
                # a ugly hack to prevent it
                else:
                    ucskip = 0
                    ignorable = False
        elif char:  # \x (not a letter)
            curskip = 0
            if char in specialchars:
                if not ignorable:
                   out += specialchars[char]
            elif char == "*":
                ignorable = True
        elif word:  # \foo
            curskip = 0
            if word in destinations:
                ignorable = True
            # http://www.biblioscape.com/rtf15_spec.htm#Heading8
            elif word == "ansicpg":
                encoding = f"cp{arg}"
                try:
                    codecs.lookup(encoding)
                except LookupError:
                    encoding = "utf8"
            if ignorable:
                pass
            elif word in specialchars:
                out += specialchars[word]
            elif word == "uc":
                ucskip = int(arg)
            elif word == "u":
                # because of https://github.com/joshy/striprtf/issues/6
                if arg is None:
                    curskip = ucskip
                else:
                    c = int(arg)
                    if c < 0:
                        c += 0x10000
                    out += chr(c)
                    curskip = ucskip
        elif _hex:  # \'xx
            if curskip > 0:
                curskip -= 1
            elif not ignorable:
                c = int(_hex, 16)
                if not hexes:
                    hexes = _hex
                else:
                    hexes += _hex
        elif tchar:
            if curskip > 0:
                curskip -= 1
            elif not ignorable:
                out += tchar
    return out

from html.parser import HTMLParser

class MyHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.items = []  # List to store text and href in order

    def handle_starttag(self, tag, attrs):
        # Check if the tag is an anchor tag
        if tag == 'a':
            # Extract the href attribute
            for attr in attrs:
                if attr[0] == 'href':
                    self.items.append(attr[1])  # Append href value directly

    def handle_data(self, data):
        # Append text data directly
        if data:  # Only append non-empty text
            self.items.append(data)

    def get_text(self):
        # Join all items into a single string
        return ''.join(self.items)

def extract_text_from_doc(doc_name):
    """Extract text from .doc file using win32com (Windows only)."""
    doc_path = os.path.join(os.getcwd(), doc_name)
    
    if not os.path.isfile(doc_path):
        raise FileNotFoundError(f"The file {doc_path} does not exist.")
    
    word = win32com.client.Dispatch("Word.Application")
    try:
        doc = word.Documents.Open(doc_path)
        
        # Initialize text variable
        text = ""
        # Iterate through paragraphs to capture all text
        for para in doc.Paragraphs:
            text += para.Range.Text + "\n"  # Add a newline for separation
        
        # Close the document without saving
        doc.Close(False)
        
    finally:
        # Quit Word application
        word.Quit()
    
    return text.strip()  # Strip any extra newlines at the end

def extract_text_from_docx(file_name):
    file_path = f'./{file_name}'
    """Extract text from .docx file."""
    doc = Document(file_path)
    text = []
    for para in doc.paragraphs:
        text.append(para.text)
    return '\n'.join(text)

def list_files_in_current_directory():
    """List all files in the current directory."""
    return [f for f in os.listdir(".") if os.path.isfile(os.path.join(".", f))]

def extract_date_and_time_from_line(line):
    # Adjust the regex pattern to handle irregular spacing in date and time
    pattern = re.compile(r'(\d{2}/\d{2}/\s*\d{4})\s*[^0-9]*\s*(\d{1,2}:\d{2}(?::\d{2})?)')
    match = pattern.search(line.replace(' ', ''))
    if match:
        date = match.group(1).replace(' ', '')  # Remove any extra spaces from the date
        time = match.group(2).replace(' ', '')
    else:
        date = None
        time = None
    return date, time

def write_to_excel(filename, headers, rows):
    """Write rows to an Excel file, appending to existing data if the file exists."""
    
    if not filename.endswith('.xlsx'):
        raise ValueError("The filename must end with '.xlsx'")

    # Create a new DataFrame with the new rows
    df_new = pd.DataFrame(rows, columns=headers)
    
    try:
        if os.path.isfile(filename):
            # Try reading the existing Excel file
            with pd.ExcelFile(filename, engine='openpyxl') as xl:
                df_existing = xl.parse(sheet_name=0)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            
            # Write the combined DataFrame to the Excel file
            with pd.ExcelWriter(filename, engine='xlsxwriter', engine_kwargs={'options':{'strings_to_formulas': False}}) as writer:
                df_combined.to_excel(writer, index=False, sheet_name='Sheet1')
        else:
            # Write the new DataFrame to a new Excel file
            with pd.ExcelWriter(filename, engine='xlsxwriter', engine_kwargs={'options':{'strings_to_formulas': False}}) as writer:
                df_new.to_excel(writer, index=False, sheet_name='Sheet1')
    except Exception as e:
        print(f"An error occurred: {e}")
        

def extract_pin(line):
    splitted_line = line.split("-")
    if len(splitted_line) > 1:
        return splitted_line[-1].split("_")[0].strip()
    else: 
        return line

def update_last_list_nth(list_of_lists, position, value):
    if list_of_lists:
        last_list = list_of_lists[-1]
        if last_list[position] == "":
            last_list[position] = value

def add_interceptacoes_from_lines_telefonica(output_files, interceptacoes, lines, file_name, pagina = None):
    next_line_is_transcricao = False
    id_num = 0
    
    for line in lines:    
        stripped = list(map(str.strip, line.split(':', 1)))
        if len(stripped) > 1:
            if stripped[0] == "file":
                stripped = list(map(str.strip, stripped[1].split(':', 1)))

            match stripped[0].replace(" ", "").lower():
                case x if "ind" in x or "índ" in x: 
                    id = stripped[1]
                    if len(interceptacoes) > 1000000:
                        output_files.append(interceptacoes[:])
                        interceptacoes.clear()
                        if pagina:
                            id_num = pagina
                        else:
                            id_num = id_num + 1
                    operacao = nome_alvo = fone_alvo = fone_contato = data = hora = duracao = tipo = direcao = obs =transcricao = attachment_name = assunto =""
                    interceptacao = [id, operacao, nome_alvo, fone_alvo, fone_contato, data, hora, duracao, tipo, direcao, obs, transcricao, attachment_name, assunto, file_name, id_num]
                    interceptacoes.append(interceptacao)
                    next_line_is_transcricao = False
                case x if "operação" in x: 
                    update_last_list_nth(interceptacoes, 1, stripped[1])
                    next_line_is_transcricao = False  
                case x if "nome" in x and "alvo" in x: 
                    update_last_list_nth(interceptacoes, 2, stripped[1])
                    next_line_is_transcricao = False 
                case x if "fone" in x and "alvo" in x:
                    update_last_list_nth(interceptacoes, 3, stripped[1])
                    next_line_is_transcricao = False 
                case x if "fone" in x and "contato" in x:
                    update_last_list_nth(interceptacoes, 4, stripped[1])
                    next_line_is_transcricao = False 
                case x if "data" in x:
                    update_last_list_nth(interceptacoes, 5, stripped[1])
                    next_line_is_transcricao = False 
                case x if "horário" in x: 
                    update_last_list_nth(interceptacoes, 6, stripped[1])
                    next_line_is_transcricao = False 
                case x if "observa" in x:
                    update_last_list_nth(interceptacoes, 9, stripped[1])
                    next_line_is_transcricao = False 
                case x if "transcri" in x:
                    update_last_list_nth(interceptacoes, 10, stripped[1])
                    next_line_is_transcricao = True
        if next_line_is_transcricao and "transcri" not in stripped[0].replace(" ", "").lower() :
                if interceptacoes:
                    last_list = interceptacoes[-1]
                    last_list[8] = last_list[8] + "\n" + line.replace('\x00', '')

def add_interceptacoes_from_txt_lines_telefonica(output_files, interceptacoes, lines, file_name, pattern):
    match_num = 0
    for line in lines:
        match = re.search(pattern, line)
        if match:
            match_num = match_num + 1
            id = match_num
            data = match.group('data').strip().replace('\x01', '')
            hora = match.group('hora').strip().replace('\x01', '')
            duracao = match.group('duracao').strip().replace('\x01', '')
            fone_alvo = match.group('fonealvo').strip().replace('\x01', '')
            attachment_name = match.group('arquivo').strip().replace('\x01', '')
            if len(interceptacoes) > 1000000:
                output_files.append(interceptacoes[:])
                interceptacoes.clear()
            interceptacao = [id, "", "", fone_alvo, "", data, hora, duracao, "", "", "", "", attachment_name, "", file_name, match_num]
            interceptacoes.append(interceptacao)

def add_interceptacoes_from_pdf_lines_telefonica(output_files, interceptacoes, lines, file_name, pagina, operacao):
    next_line_is_transcricao = False
    pattern_interceptado = r'(\d{2}:\d{2}:\d{2})(.*?)Nº Interceptado: (.*)'
    pattern_assunto_fone_contato = r'(.*)Assunto: Nº Contato:(.*)'
    for line in lines:      
        match line:
            case x if x.startswith("Duração:"): 
                id = x.split('Duração:', 1)[1].strip()
                if len(interceptacoes) > 1000000:
                    output_files.append(interceptacoes[:])
                    interceptacoes.clear()
                nome_alvo = fone_alvo = fone_contato = data = hora = duracao = tipo = direcao = obs = transcricao = attachment_name = assunto =""
                interceptacao = [id, operacao, nome_alvo, fone_alvo, fone_contato, data, hora, duracao, tipo, direcao, obs, transcricao, attachment_name, assunto, file_name, pagina]
                interceptacoes.append(interceptacao)
                next_line_is_transcricao = False
            case x if "Nº Interceptado:" in x:
                match = re.search(pattern_interceptado, line)
                fone_alvo = duracao = nome_alvo = ""
                if match:
                    duracao = match.group(1).strip()
                    nome_alvo = match.group(2).strip()
                    fone_alvo = match.group(3)
                else:
                    nome_alvo = x.split('Nº Interceptado:', 1)[0].strip()
                    if nome_alvo.startswith('-'):
                        nome_alvo = nome_alvo[1:]
                    fone_alvo = x.split('Nº Interceptado:', 1)[1].strip()
                update_last_list_nth(interceptacoes, 2, nome_alvo)
                update_last_list_nth(interceptacoes, 3, fone_alvo)
                update_last_list_nth(interceptacoes, 7, duracao)
                next_line_is_transcricao = False
            case x if "Tipo:Direção:" in x:
                splitted = x.split("Tipo:Direção:")
                tipo = splitted[0].strip()
                direcao = splitted[1].strip()
                update_last_list_nth(interceptacoes, 8, tipo)
                update_last_list_nth(interceptacoes, 9, direcao)
                next_line_is_transcricao = False
            case x if "Data:" in x: 
                date, time = extract_date_and_time_from_line(line)
                update_last_list_nth(interceptacoes, 5, date)
                update_last_list_nth(interceptacoes, 6, time)
                next_line_is_transcricao = False
            case x if x.startswith("Arquivo:"): 
                attachment_name = x.split('Arquivo:', 1)[1].strip()
                update_last_list_nth(interceptacoes, 11, attachment_name)
                next_line_is_transcricao = False
            case x if "Assunto: Nº Contato:" in x:
                splitted = x.split("Assunto: Nº Contato:", 1)
                assunto = splitted[0].strip()
                fone_contato = splitted[1].strip()
                update_last_list_nth(interceptacoes, 4, fone_contato)
                update_last_list_nth(interceptacoes, 12, assunto)
                next_line_is_transcricao = False
            case x if "Degravação:" in x:
                next_line_is_transcricao = True
            case x if "TAGs Atribuídas:" in x:
                next_line_is_transcricao = False
        if next_line_is_transcricao and all(substring not in line for substring in ["Operação", "SIS", "Página", "Degravação"]):
            if interceptacoes:
                    last_list = interceptacoes[-1]
                    last_list[10] = last_list[10] + "\n" + line.replace('\x00', '')
            
def add_interceptacoes_from_html_content_telefonica (output_files, interceptacoes, content, file_name):
    pattern_relatorio = re.compile(r'Mídia\s*Interlocutor\s*Data\/Hora\s*Inicial\s*Duração\s*Dados\s*[\s\S]*Comentário\s*Nome\s*do\s*Arquivo')
    pattern_chamada = re.compile(r'Interlocutor\s*Data\/Hora\s*Inicial\s*Duração\s*[\s\S]*Comentário\s*Nome\s*do\s*Arquivo')
    pattern_index = re.compile(r'TELEFONE\s*INTERLOCUTOR\s*DATA\/HORA\s*INICIAL\s*DURAÇÃO\s*DADOS\s*ÁUDIO\s*INTERLOCUTORES\/COMENTÁRIO')
    pattern_relatorio_2 = re.compile(r'ID\s*Telefone do Alvo\s*Telefone\s*do\s*Contato\s*Tipo\s*Direção\s*Data\s*Duração\s*Assunto\s*Comunicação\s*Degravação\s*TAGs\s*Atribuídas')
    pattern_dados_gravacao = re.compile(r'TELEFONE\s*INTERLOCUTOR\s*DATA\/HORA\s*INICIAL\s*DATA\/HORA\s*FINAL\s*DURAÇÃO\s*ÁUDIO\s*INTERLOCUTORES\/COMENTÁRIO')
    pattern_relatorio_3 = re.compile(r""" 
        \s*(?P<codigo>.+)\s*
        Data:\s*(?P<data>.+?)\s*Hora:\s*(?P<hora>.+?)\s*Duração:\s*(?P<duracao>.+?)\s*
        Alvo:\s*(?P<alvo>.+?)\s*
        Fone\s+Alvo:\s*(?P<fonealvo>.+?)\s*Fone\s+Contato:\s*(?P<fonecontato>.+?)\s*
        Interlocutores:\s*(?P<interlocutor>.+?)\s*
        (?:Arquivo:)?\s*(?P<arquivo>.+?)\s*
        Degravação:\s*(?P<degravacao>.+?)
    """, re.VERBOSE | re.DOTALL)
    if pattern_relatorio.search(content):
        pattern = r'\r?\n\s*\r?\n\s*\d+\r?\n\s*\r?\n'
        blocks = re.split(pattern, content)
        block_num = 0
        for block in blocks[1:]:
            block_num = block_num + 1
            lines = block.split("\n")
            fone_alvo = lines[0]
            fone_contato = lines[1]
            data = lines[2].split(" ")[0]
            hora = data = lines[2].split(" ")[1]
            duracao = lines[3]
            id = lines[4].split(".",1)[0].split("CHAMADA_")[1]
            obs = lines[6]
            if len(lines) > 7:
                attachment_name = lines[7]
            else:
                attachment_name = ""
            if len(interceptacoes) > 1000000:
                output_files.append(interceptacoes[:])
                interceptacoes.clear()
            interceptacao = [id, "", "", fone_alvo, fone_contato, data, hora, duracao, "", "", obs, "", attachment_name, "", file_name, block_num]
            interceptacoes.append(interceptacao)
    elif pattern_chamada.search(content):
        pattern = r"""
            (?s)
            Mídia\s*
            Nome\s+do\s+Alvo\s*
            (?P<phone>\d{2}\(\d{2}\)\d{8,9})\s*
            (?P<name1>[^\n]+?)\s*
            Nome\s+do\s+Alvo:\s*(?P<name2>[^\n]+?)\s*
            Relação\s+das\s+Transcrições\s*
            (?:.*?\s*)?
            Interlocutor\s*
            Data\/Hora\s+Inicial\s*
            Duração\s*
            [\s\S]*?
            (?: (?P<comentario>Comentário\s+[^\n]*)\s* )?
            (?: Nome\s+do\s+Arquivo\s* (?P<arquivo>[^\s]+?)\s* )?
            (?P<val>\d*)\s*
            (?P<phone2>\d*)\s*
            (?P<date>\d{2}\/\d{2}\/\d{4})\s+(?P<time>\d{2}:\d{2}:\d{2})\s*
            (?P<duration>\d{2}:\d{2}:\d{2})\s*
            (?P<y>[^\n]*)\s*
            (?P<description>[^\n]*)\s*
            (?P<attachment_name>.+)\s*
            Transcrição\s*(?P<transcription>.+)
            """

        matches = re.search(pattern, content, re.VERBOSE)
        if matches:
            fone_alvo = matches.group('phone')
            nome_alvo = matches.group('name1')
            if not nome_alvo:
                nome_alvo = matches.group('name2')
            fone_contato = matches.group('phone2')
            data = matches.group('date')
            hora = matches.group('time')
            duracao = matches.group('duration')
            obs = matches.group('description')
            attachment_name = matches.group('attachment_name').split('\n')[0]
            id = file_name.split(".",1)[0].split("CHAMADA_")[1]
            transcricao = matches.group('transcription')
            if len(interceptacoes) > 1000000:
                output_files.append(interceptacoes[:])
                interceptacoes.clear()
            interceptacao = [id, "", nome_alvo, fone_alvo, fone_contato, data, hora, duracao, "", "", obs, transcricao, attachment_name, "", file_name, 1]
            interceptacoes.append(interceptacao)
            
        else:
            print("Não foi possível extrair os dados desse arquivo")
    elif pattern_index.search(content):
        pattern = r'(?s)(?P<phone1>\d{10,11})(?:\s+(?P<phone2>\d{10,11}))?\s+(?P<date>\d{2}\/\d{2}\/\d{4})\s+(?P<time>\d{2}:\d{2}:\d{2})\s+(?P<duration>\d{2}:\d{2}:\d{2})\s+(?P<transcript_path>Transcricoes\/\d+_\d{14}_\d{3}_\d+\.html)\s+(?P<attachment_name>Gravacoes\/\d+_\d{14}_\d{3}_\d+\.wav)\s*(?P<description>[^\xa0]+)'
        compiled_pattern = re.compile(pattern, re.DOTALL)
        matches = re.finditer(compiled_pattern, content)
        match_num=0
        for match in matches:
            match_num = match_num + 1
            fone_alvo = match.group('phone1')
            fone_contato = match.group('phone2')
            data = match.group('date')
            hora = match.group('time')
            duracao = match.group('duration')
            obs = match.group('description')
            attachment_name = match.group('attachment_name').split('/',1)[-1]
            id = attachment_name.split(".",1)[0].split("_")[-1]
            if len(interceptacoes) > 1000000:
                output_files.append(interceptacoes[:])
                interceptacoes.clear()
            interceptacao = [id, "", "", fone_alvo, fone_contato, data, hora, duracao, "", "", obs, "", attachment_name, "", file_name, match_num]
            interceptacoes.append(interceptacao)
        
    elif pattern_dados_gravacao.search(content):
        pattern = r'(?s)TELEFONE.*ALVO\s+(?P<phone>\d+)\s+(?P<name>.+).*TELEFONE.*COMENTÁRIO[\s\xA0]*(?P<phone2>\d+)[\s\xA0]*(?P<date>[\S]+)[\s\xA0]*(?P<time>[\S]+)[\s\xA0]*(?P<x>[\S]+)[\s\xA0]*(?P<y>[\S]+)[\s\xA0]*(?P<duration>[\S]+)[\s\xA0]*(?P<attachment_name>[\S]+)[\s\xA0]*(?P<description>.*)'
        matches = re.search(pattern, content, re.VERBOSE)
        if matches:
            fone_alvo = matches.group('phone').strip()
            nome_alvo = matches.group('name').strip()
            if not fone_alvo:
                fone_alvo = matches.group('phone2').strip()
            data = matches.group('date').strip()
            hora = matches.group('time').strip()
            duracao = matches.group('duration').strip()
            obs = matches.group('description').strip()
            attachment_name = matches.group('attachment_name').strip().split("/")[-1]
            id = file_name.split(".",1)[0].split("_")[-1]
            if len(interceptacoes) > 1000000:
                output_files.append(interceptacoes[:])
                interceptacoes.clear()
            interceptacao = [id, "", nome_alvo, fone_alvo, "", data, hora, duracao, "", "", obs, "", attachment_name, "", file_name, 1]
            interceptacoes.append(interceptacao)
    
    elif pattern_relatorio_2.search(content):
        #pattern = r'(?s)(?P<id>\d{8,20})\s*(?P<phone1>\(\d{2}\)\d{4,5}\-\d{4})\s*(?P<phone2>[^\n]*)\s*(?P<tipo>SMS|Áudio)\s*(?P<date>\d{2}\/\d{2}\/\d{4})\s*(?P<time>\d{2}:\d{2}:\d{2})\s*(?P<duration>[^\n]*)\s*(?P<description>[^\n]*)\s*(?P<attachment_name>[^\n]*)\s*(?P<transcription>[^\xa0]+)'
        pattern = r'(?s)(?P<id>\d{8,20})\s*(?P<phone1>\(\d{2}\)\d{4,5}\-\d{4}|\d{11,20})\s*(?P<phone2>[^\n]*)\s*(?P<tipo>SMS|Áudio)\s*(?P<direction>[^\n]*)\s*(?P<date>\d{2}\/\d{2}\/\d{4})\s*(?P<time>\d{2}:\d{2}:\d{2})\s*(?P<duration>[^\n]*)\s*(?P<description>[^\n]*)\s*(?P<attachment_name>[^\n]*)'
        compiled_pattern = re.compile(pattern, re.DOTALL)
        matches = list(re.finditer(compiled_pattern, content))  # Convert to list to make it indexable
        match_num = 0
        last_end = 0

        for i, match in enumerate(matches):
            match_num += 1
            fone_alvo = match.group('phone1')
            fone_contato = match.group('phone2')
            data = match.group('date')
            hora = match.group('time')
            duracao = match.group('duration')
            tipo = match.group('tipo')
            direction = match.group('direction')
            obs = match.group('description')
            attachment_name = match.group('attachment_name').split('/',1)[-1]
            id = match.group('id')
            
            between_content = content[last_end:match.start()].strip() if i > 0 else ""
            last_end = match.end()
            
            transcription = ""
            if "Texto degravação:" in between_content:
                transcription = between_content.split("Texto degravação:", 1)[1].strip()

            if len(interceptacoes) > 1000000:
                output_files.append(interceptacoes[:])
                interceptacoes.clear()
            
            interceptacao = [id, "", "", fone_alvo, fone_contato, data, hora, 
                            duracao, tipo, direction, obs, transcription, attachment_name, "", 
                            file_name, match_num] 
            interceptacoes.append(interceptacao)
    elif pattern_relatorio_3.search(content):
        pattern_operacao = r"""
            Operação:(?P<operacao>[^\n\r]*)
            """
        match_operacao = re.search(pattern_operacao, content, re.VERBOSE)
        if match_operacao:
            operacao = match_operacao.group('operacao').replace("\r", "").strip().replace('\x01', '').replace("\n", "")
        else:
            operacao = ""
        pattern = r""" 
            \s*(?P<codigo>.*)\s*
            Data:\s*(?P<data>.+?)\s*Hora:\s*(?P<hora>.+?)\s*Duração:\s*(?P<duracao>.+?)\s*
            Alvo:\s*(?P<alvo>.+?)\s*
            Fone\s+Alvo:\s*(?P<fonealvo>.+?)\s*Fone\s+Contato:\s*(?P<fonecontato>.+?)\s*
            Interlocutores:\s*(?P<interlocutor>.*)
        """
        blocks = re.split(r'Código:', content)
        match_num = 0
        operacao_next = operacao
        for block in blocks:
            block = re.sub(r'^[a-fA-F0-9]{1,}\n', '', block, flags=re.MULTILINE)
            match = re.search(pattern, block, re.VERBOSE | re.DOTALL)
            if match:
                match_num = match_num + 1
                id = match.group('codigo').strip().replace('\x01', '').replace('@TAB', '')
                data = match.group('data').strip().replace('\x01', '').replace('@TAB', '')
                hora = match.group('hora').strip().replace('\x01', '').replace('@TAB', '')
                duracao = match.group('duracao').strip().replace('\x01', '').replace('@TAB', '')
                nome_alvo = match.group('alvo').strip().replace('\x01', '').replace('@TAB', '')
                fone_alvo = match.group('fonealvo').strip().replace('\x01', '').replace('@TAB', '')
                fone_contato = match.group('fonecontato').strip().replace('\x01', '').replace('@TAB', '')
                interlocutor = match.group('interlocutor').strip().replace('\x01', '').replace('@TAB', '')
                # Extrair nome do arquivo .wav dentro do texto de interlocutor
                attachment_match = re.search(r"\d{17}\.wav", interlocutor)
                attachment_name = attachment_match.group(0) if attachment_match else ""

                # Separar tudo depois de "Degravação:"
                full_text = match.group(0)  # o texto completo do match (interlocutor + .wav + Degravação)
                split = re.split(r'Degravação:*', full_text, maxsplit=1, flags=re.IGNORECASE)
                degravacao = split[1].strip() if len(split) > 1 else ""
                
                match_operacao = re.search(pattern_operacao, degravacao, re.VERBOSE)
                if match_operacao:
                    operacao_next = match_operacao.group('operacao').replace("\r", "").strip().replace('\x01', '')
                if len(interceptacoes) > 1000000:
                    output_files.append(interceptacoes[:])
                    interceptacoes.clear()
                interceptacao = [id, operacao, nome_alvo, fone_alvo, fone_contato, data, hora, duracao, "", "", "", degravacao, attachment_name, "", file_name, match_num]
                interceptacoes.append(interceptacao)
                operacao = operacao_next
    else:
        print("Não foi possível identificar o padrão do arquivo, este foi ignorado.")

def main():
    try:
        print("Iniciando o programa!")
        field_names = ["Índice", "Operação", "Nome do Alvo", "Fone do Alvo", "Fone de Contato", "Data", "Horário", "Duração", " Tipo", "Direção", "Observação", "Transcrição", "Arquivo anexado", "Assunto", "Arquivo", "Página/Posição"]
        id = pacote = date = time = direcao = alvo = contato = perfil = mensagem = ""
        file_num = 1
        interceptacoes=[]
        output_files=[]
        content=''
        interceptacoes_antigo = 0
        for file_name in list_files_in_current_directory():
            next_line_is_transcricao = False
            _, ext = os.path.splitext(file_name)
            if ext.lower() not in ('.py', '.ipynb', '.cif', '.jar', '.exe', '.xlsx', 'xls'):
                print(str(file_num) + ") Extraindo capturas do arquivo: " + file_name)
                file_num = file_num + 1
                pattern_relatorio_3 = re.compile(
                    r"Código:.*?Data:.*?Hora:.*?Duração:.*?Alvo:.*?Fone\s+Alvo:.*?Fone\s+Contato:.*?Interlocutores:.*?(?:Arquivo:.*?)?Degravação:",
                    re.DOTALL | re.IGNORECASE
                )
                pattern_txt_2 = re.compile(r"""\s*(?P<arquivo>.*)\s*---\sFONE\s+ALVO:\s*(?P<fonealvo>.+?)\s*-\s*Data:\s*(?P<data>.+?)\s*-\s*Hora:\s*(?P<hora>.+?)\s*-\s*Duração:\s*(?P<duracao>.*)\s*""", re.VERBOSE | re.DOTALL)
                pattern_microsoft = re.compile(r"microsoft\s*windows|Windows\s*PowerShell|Windows\s*Server|Microsoft\s*Corporation", re.IGNORECASE)
                if ext.lower() in ('.html', '.htm', '.txt', '.rtf'):
                    with open(file_name, 'r', encoding="latin1") as file:
                        content = file.read()
                        try:
                            content = content.encode('latin1').decode('utf-8')
                        except UnicodeDecodeError:
                            pass  # fallback if it fails, keep as-is or handle differently
                        if ext.lower() in ('.html', '.htm'):
                            parser = MyHTMLParser()
                            parser.feed(content)
                            content = parser.get_text()
                            add_interceptacoes_from_html_content_telefonica(output_files, interceptacoes, content, file_name)
                        # if ext.lower() in ('.rtf') and pattern_microsoft.search(content):
                        #     pass
                        else:
                            lines = content.split("\n")
                            if ext.lower() in ('.rtf'):
                                lines = [format_rtf_line(line) for line in lines]
                            content = "\n".join(lines)
                            
                            if pattern_txt_2.search(content[:5000]):
                                add_interceptacoes_from_txt_lines_telefonica(output_files, interceptacoes, lines, file_name, pattern_txt_2)
                            elif pattern_relatorio_3.search(content):
                                add_interceptacoes_from_html_content_telefonica(output_files, interceptacoes, content, file_name)
                            else:
                                add_interceptacoes_from_lines_telefonica(output_files, interceptacoes, lines, file_name)
                elif ext.lower() == '.docx':
                    content = extract_text_from_docx(file_name)
                    if pattern_relatorio_3.search(content[:5000]):
                        add_interceptacoes_from_html_content_telefonica(output_files, interceptacoes, content, file_name)
                    else:
                        lines = content.split("\n")
                        add_interceptacoes_from_lines_telefonica(output_files, interceptacoes, lines, file_name)
                elif ext.lower() == '.doc':
                    content = extract_text_from_doc(file_name)
                    if pattern_relatorio_3.search(content[:5000]):
                        add_interceptacoes_from_html_content_telefonica(output_files, interceptacoes, content, file_name)
                    else:
                        lines = content.split("\n")
                        add_interceptacoes_from_lines_telefonica(output_files, interceptacoes, lines, file_name)
                elif ext.lower() == '.pdf':
                    with open(file_name, 'rb') as file:
                        reader = PdfReader(file)
                        page_num=0
                        operacao = ""
                        for page in reader.pages:
                            lines = page.extract_text().split("\n")
                            for line in lines:      
                                match line:
                                    case x if x.startswith("Operação"):             
                                        operacao = x.split("Operação", 1)[1].strip()
                            break
                        for page in reader.pages:
                            page_num = page_num + 1
                            lines = page.extract_text().split("\n")
                            add_interceptacoes_from_pdf_lines_telefonica(output_files, interceptacoes, lines, file_name, page_num, operacao)
                            
                            
                else:
                    print("Arquivo não compatível: " + file_name)
                print(f"Total de {sum(len(file) for file in output_files) + len(interceptacoes)} linhas até aqui.")
                codigos = content.count("Código:")
                if codigos == 0:
                    codigos = content.count("FONE ALVO:")
                file_interceptacoes = len(interceptacoes) - interceptacoes_antigo
                interceptacoes_antigo = len(interceptacoes)
                if codigos != file_interceptacoes:
                    print(f"Total de {codigos} Código: ou FONE ALVO: nesse arquivo e {file_interceptacoes} inteceptações nesse arquivo.")
        output_files.append(interceptacoes[:])
        print("Total de " + str(sum(len(file) for file in output_files)) +" somando todos os arquivos.")

        output_file_num = 0
        if len(output_files) == 1:
            print("Escrevendo o arquivo base_de_dados_TELEFONICA.xlsx")
            write_to_excel("base_de_dados_TELEFONICA.xlsx", field_names, output_files[0])
            print("Arquivo base_de_dados_BBM.xlsx escrito com sucesso")
        else: 
            for output_file in output_files:
                output_file_num = output_file_num + 1
                print("Escrevendo o arquivo base_de_dados_TELEFONICA.xlsx parte " + str(output_file_num) + " de " + str(len(output_files)))
                write_to_excel("base_de_dados_TELEFONICA_PART_" + str(output_file_num) + ".xlsx", field_names, output_file)
                print("Arquivo base_de_dados_TELEFONICA.xlsx parte " + str(output_file_num) + " de " + str(len(output_files)) + " escrito com sucesso")

        print("Processo finalizado com sucesso!!")
    except Exception as e:
        print(f"An error occurred: {e}")
main()
print("The program has finished. Press Enter to exit.")
input()