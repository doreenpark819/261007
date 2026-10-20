"""VBA 소스(UTF-8)를 ASCII 전용으로 변환: 한글 문자열 → T("\\uXXXX"), 한글 주석 제거, 한글 Const → 함수."""
import re, sys

src = open(sys.argv[1], encoding="utf-8").read().splitlines()
T_FUNC = '''
' Decode "\\uXXXX" escapes (keeps this file ASCII-only so encoding never breaks)
Private Function T(ByVal s As String) As String
    Dim i As Long, o As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 2) = "\\u" Then
            o = o & ChrW(CLng("&H" & Mid$(s, i + 2, 4) & "&"))
            i = i + 6
        Else
            o = o & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    T = o
End Function
'''.strip("\n").splitlines()


def esc(s):
    return "".join(c if ord(c) < 128 else f"\\u{ord(c):04X}" for c in s)


def split_code_comment(line):
    """(code, comment) — comment starts at ' outside string literal."""
    ins = False
    for i, ch in enumerate(line):
        if ch == '"':
            ins = not ins
        elif ch == "'" and not ins:
            return line[:i], line[i:]
    return line, ""


def conv_strings(code):
    out, i = [], 0
    while i < len(code):
        if code[i] == '"':
            j = i + 1
            while True:
                if code[j] == '"' and j + 1 < len(code) and code[j + 1] == '"':
                    j += 2
                    continue
                if code[j] == '"':
                    break
                j += 1
            lit = code[i:j + 1]
            body = lit[1:-1]
            out.append(f'T("{esc(body)}")' if any(ord(c) > 127 for c in body) else lit)
            i = j + 1
        else:
            out.append(code[i])
            i += 1
    return "".join(out)


res, consts = [], []
for line in src:
    m = re.match(r'^Private Const (\w+) As String = "(.*)"\s*$', line)
    if m and any(ord(c) > 127 for c in m.group(2)):
        consts += [f"Private Function {m.group(1)}() As String", f'    {m.group(1)} = T("{esc(m.group(2))}")', "End Function", ""]
        continue
    code, comment = split_code_comment(line)
    if any(ord(c) > 127 for c in comment):
        comment = ""
    code = conv_strings(code)
    new = (code + comment).rstrip()
    if not new.strip() and line.strip():
        continue                        # 한글 주석만 있던 줄은 삭제
    res.append(new)

# Option Explicit 및 모듈 변수 선언 뒤에 T()와 상수 함수 삽입
k = max(i for i, l in enumerate(res) if l.startswith("Private wd As Object")) + 1
res = res[:k] + [""] + T_FUNC + [""] + consts + res[k:]
text = "\r\n".join(res) + "\r\n"
assert all(ord(c) < 128 for c in text), "non-ascii remains"
open(sys.argv[2], "w", encoding="ascii", newline="").write(text)
print(len(res), "lines")
