// ============================================================================
// APPS SCRIPT — CERTIFICADOS (.pfx) NO DRIVE  (UM SÓ PARA OS 6 ESCRITÓRIOS)
// ============================================================================
// Guarda o ARQUIVO .pfx/.p12 que o escritório importa no menu CERTIFICADO
// DIGITAL do *_Gestor_Fiscal.py numa pasta do Drive do dono deste script:
//
//   CERTIFICADOS/                 ← pasta criada à mão no Drive
//   ├── ASBEM/  *.pfx + senhas.json
//   ├── VIDAL/  *.pfx + senhas.json
//   ├── VS/ WR/ VISAO/ VJ/
//
// senhas.json de cada subpasta = {"<nome do arquivo>": "<senha>"}, MESMO
// formato de escritorios/<X>/certificados/senhas.json do SISTEMA BUSCA
// NACIONAL — dá pra copiar a subpasta inteira pra lá.
//
// Diferente de apps_script_certificado_digital.gs (um por planilha), este é
// STANDALONE e COMPARTILHADO, igual ao apps_script_pdf_situacao_fiscal.gs:
// o escritório vem no corpo do POST. Não fica anexado a planilha nenhuma.
//
// SEGURANÇA — .pfx + senha = assinatura digital completa da empresa:
//   • A pasta CERTIFICADOS NÃO pode ser compartilhada por link.
//   • ENVIAR não pede token (decisão do usuário, 26/09: a URL fica no código
//     do Gestor, sem precisar mexer nos Secrets do Streamlit). Quem tiver a
//     URL consegue só COLOCAR arquivo na pasta — nunca ver, baixar, apagar
//     nem sobrescrever (arquivo de mesmo nome e conteúdo diferente é
//     guardado como "nome (2).pfx").
//   • Listar/baixar exige o TOKEN_ADMIN (só pro SISTEMA BUSCA NACIONAL),
//     gerado pela função `configurar` e guardado nas Propriedades do Script.
//
// PUBLICAR (uma vez só, na conta dona da pasta CERTIFICADOS):
//   1. script.google.com > Novo projeto > colar este código > Salvar.
//   2. Implantar > Nova implantação > "Aplicativo da Web".
//      Executar como: Eu. Quem pode acessar: Qualquer pessoa.
//   3. Colar a URL em CERT_DRIVE_URL no *_Gestor_Fiscal.py de cada escritório.
//   A pasta CERTIFICADOS é achada sozinha no primeiro envio (pelo nome, em
//   qualquer lugar do Drive do dono — precisa existir só UMA com esse nome).
//   `configurar` (opcional) cria as 6 subpastas de uma vez e mostra o
//   TOKEN_ADMIN — só vai ser preciso quando a Busca Nacional for ler daqui.
//
// Mudou o código? Colar aqui de novo (este arquivo é só backup) e
// "Implantar > Gerenciar implantações > lápis > Nova versão" pra manter a URL.
// ============================================================================

const ESCRITORIOS = ["ASBEM", "VIDAL", "VS", "WR", "VISAO", "VJ"];
const NOME_PASTA_RAIZ = "CERTIFICADOS";
const NOME_SENHAS = "senhas.json";
const TAMANHO_MAX = 1024 * 1024;   // 1 MB — um .pfx tem poucos KB


// ── Rodar à mão pelo editor (opcional) ───────────────────────────────────────
function configurar() {
  const props = PropertiesService.getScriptProperties();
  const raiz = pastaRaiz_();
  if (raiz.getSharingAccess() !== DriveApp.Access.PRIVATE) {
    Logger.log("⚠️ ATENÇÃO: a pasta CERTIFICADOS está compartilhada por link. Deixe-a restrita.");
  }
  ESCRITORIOS.forEach(function (esc) { subpasta_(esc); });

  if (!props.getProperty("TOKEN_ADMIN")) {
    props.setProperty("TOKEN_ADMIN", Utilities.getUuid().replace(/-/g, ""));
  }
  Logger.log("Pasta: " + raiz.getUrl() + "\nSubpastas: " + ESCRITORIOS.join(", ") +
             "\n\nTOKEN_ADMIN (só pra Busca Nacional): " + props.getProperty("TOKEN_ADMIN"));
}


// ── Web app ──────────────────────────────────────────────────────────────────
function doGet() {
  // Só pra conferir se a implantação está no ar — não devolve dado nenhum.
  return json_({ status: "ok", servico: "certificados-drive", versao: 3 });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const acao = req.acao || "upload";
    if (acao === "upload") return json_(upload_(req));
    if (acao === "upload_lote") return json_(uploadLote_(req));
    if (acao === "listar") return json_(listar_(req));
    if (acao === "baixar") return json_(baixar_(req));
    return json_({ status: "erro", mensagem: "Ação desconhecida: " + acao });
  } catch (err) {
    return json_({ status: "erro", mensagem: String(err && err.message || err) });
  }
}


// ── upload (1 arquivo): {escritorio, nome_arquivo, conteudo_b64, senha,
//                          cnpj, razao_social, validade_iso} ────────────────
function upload_(req) {
  const r = uploadLote_({ escritorio: req.escritorio, itens: [req] });
  if (r.status !== "ok") return r;
  return r.itens[0];
}


// ── upload_lote (vários numa chamada só — bem mais rápido que um POST por
//    arquivo: a pasta e o senhas.json são lidos/gravados UMA vez):
//    {escritorio, itens: [{nome_arquivo, conteudo_b64, senha, cnpj,
//                          razao_social, validade_iso}, ...]}
//    → {status: "ok", itens: [{status, resultado|mensagem, arquivo}, ...]}
//    (um item com erro não impede os outros) ────────────────────────────────
function uploadLote_(req) {
  const esc = String(req.escritorio || "").toUpperCase();
  if (ESCRITORIOS.indexOf(esc) < 0) return { status: "erro", mensagem: "Escritório inválido: " + esc };
  const itens = req.itens || [];

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);   // dois envios do mesmo escritório mexem no mesmo senhas.json
  try {
    const pasta = subpasta_(esc);
    const senhas = lerSenhas_(pasta);
    const saida = itens.map(function (it) {
      try {
        return guardarArquivo_(pasta, senhas, it);
      } catch (err) {
        return { status: "erro", arquivo: String(it.nome_arquivo || ""), mensagem: String(err && err.message || err) };
      }
    });
    if (saida.some(function (r) { return r.status === "ok"; })) gravarSenhas_(pasta, senhas);
    return { status: "ok", itens: saida };
  } finally {
    lock.releaseLock();
  }
}


function guardarArquivo_(pasta, senhas, req) {
  let nome = String(req.nome_arquivo || "").replace(/[\\\/:*?"<>|]/g, "_").trim();
  if (!/\.(pfx|p12)$/i.test(nome)) return { status: "erro", arquivo: nome, mensagem: "Arquivo precisa ser .pfx ou .p12." };

  const bytes = Utilities.base64Decode(String(req.conteudo_b64 || ""));
  if (!bytes.length || bytes.length > TAMANHO_MAX) return { status: "erro", arquivo: nome, mensagem: "Arquivo vazio ou grande demais." };

  const md5 = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes));
  let resultado = "novo";

  // Nunca sobrescreve: mesmo nome + mesmo conteúdo → já existia;
  // mesmo nome + conteúdo diferente → guarda como "nome (2).pfx", "(3)"...
  const base = nome.replace(/\.(pfx|p12)$/i, "");
  const ext = nome.slice(base.length);
  for (let n = 2; ; n++) {
    const iguais = pasta.getFilesByName(nome);
    if (!iguais.hasNext()) break;
    let mesmoConteudo = false;
    while (iguais.hasNext()) {
      const b = iguais.next().getBlob().getBytes();
      if (hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, b)) === md5) mesmoConteudo = true;
    }
    if (mesmoConteudo) { resultado = "ja_existia"; break; }
    nome = base + " (" + n + ")" + ext;
  }

  let arq = null;
  if (resultado !== "ja_existia") {
    arq = pasta.createFile(Utilities.newBlob(bytes, "application/x-pkcs12", nome));
    arq.setDescription(JSON.stringify({
      cnpj: String(req.cnpj || ""),
      razao_social: String(req.razao_social || ""),
      validade_iso: String(req.validade_iso || ""),
      enviado_em: new Date().toISOString(),
    }));
  }
  senhas[nome] = String(req.senha || "");
  return { status: "ok", resultado: resultado, arquivo: nome, id: arq ? arq.getId() : "" };
}


// ── listar (admin): {token, escritorio?} → arquivos + senhas ────────────────
function listar_(req) {
  if (!tokenOk_(req.token, "TOKEN_ADMIN")) return { status: "erro", mensagem: "Token inválido." };
  const alvo = req.escritorio ? [String(req.escritorio).toUpperCase()] : ESCRITORIOS;
  const saida = {};
  alvo.forEach(function (esc) {
    if (ESCRITORIOS.indexOf(esc) < 0) return;
    const pasta = subpasta_(esc);
    const arquivos = [];
    const it = pasta.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (!/\.(pfx|p12)$/i.test(f.getName())) continue;
      let info = {};
      try { info = JSON.parse(f.getDescription() || "{}"); } catch (e) {}
      arquivos.push({
        id: f.getId(), nome: f.getName(), tamanho: f.getSize(),
        atualizado: f.getLastUpdated().toISOString(),
        cnpj: info.cnpj || "", razao_social: info.razao_social || "", validade_iso: info.validade_iso || "",
      });
    }
    saida[esc] = { arquivos: arquivos, senhas: lerSenhas_(pasta) };
  });
  return { status: "ok", escritorios: saida };
}


// ── baixar (admin): {token, id} → conteúdo em base64 ────────────────────────
function baixar_(req) {
  if (!tokenOk_(req.token, "TOKEN_ADMIN")) return { status: "erro", mensagem: "Token inválido." };
  const f = DriveApp.getFileById(String(req.id || ""));
  // Só entrega arquivo que esteja dentro de CERTIFICADOS/<ESCRITÓRIO>/
  const raizId = pastaRaiz_().getId();
  let dentro = false;
  const pais = f.getParents();
  while (pais.hasNext()) {
    const p = pais.next();
    const avos = p.getParents();
    while (avos.hasNext()) if (avos.next().getId() === raizId) dentro = true;
  }
  if (!dentro) return { status: "erro", mensagem: "Arquivo fora da pasta CERTIFICADOS." };
  return { status: "ok", nome: f.getName(), conteudo_b64: Utilities.base64Encode(f.getBlob().getBytes()) };
}


// ── auxiliares ───────────────────────────────────────────────────────────────
function pastaRaiz_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("PASTA_ID");
  if (id) return DriveApp.getFolderById(id);

  // Primeira vez: acha a pasta CERTIFICADOS do dono do script pelo nome
  const eu = Session.getEffectiveUser().getEmail();
  const it = DriveApp.getFoldersByName(NOME_PASTA_RAIZ);
  const achadas = [];
  while (it.hasNext()) {
    const f = it.next();
    const dono = f.getOwner();
    if (!f.isTrashed() && dono && dono.getEmail() === eu) achadas.push(f);
  }
  if (achadas.length !== 1) {
    throw new Error("Esperava 1 pasta '" + NOME_PASTA_RAIZ + "' no Drive do dono do script, achei " +
                    achadas.length + ". Renomeie as outras ou grave o ID na propriedade PASTA_ID.");
  }
  props.setProperty("PASTA_ID", achadas[0].getId());
  return achadas[0];
}

function subpasta_(esc) {
  // ID da subpasta guardado nas Propriedades → não procura pelo nome toda vez
  const props = PropertiesService.getScriptProperties();
  const chave = "SUBPASTA_" + esc;
  const id = props.getProperty(chave);
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) {}   // apagada/movida → procura de novo
  }
  const raiz = pastaRaiz_();
  const it = raiz.getFoldersByName(esc);
  const pasta = it.hasNext() ? it.next() : raiz.createFolder(esc);
  props.setProperty(chave, pasta.getId());
  return pasta;
}

function lerSenhas_(pasta) {
  const it = pasta.getFilesByName(NOME_SENHAS);
  if (!it.hasNext()) return {};
  try { return JSON.parse(it.next().getBlob().getDataAsString("UTF-8") || "{}"); } catch (e) { return {}; }
}

function gravarSenhas_(pasta, senhas) {
  const txt = JSON.stringify(senhas, null, 2);
  const it = pasta.getFilesByName(NOME_SENHAS);
  if (it.hasNext()) it.next().setContent(txt);
  else pasta.createFile(NOME_SENHAS, txt, "application/json");
}

function tokenOk_(recebido, prop) {
  const esperado = PropertiesService.getScriptProperties().getProperty(prop);
  return !!esperado && String(recebido || "") === esperado;
}

function hex_(bytes) {
  return bytes.map(function (b) { return ("0" + (b & 0xff).toString(16)).slice(-2); }).join("");
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
