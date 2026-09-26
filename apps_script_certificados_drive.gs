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
//   • Cada escritório tem o SEU token (TOKEN_VIDAL, TOKEN_VJ...), que só
//     serve pra ENVIAR certificado daquele escritório — não lê nada.
//   • Listar/baixar exige o TOKEN_ADMIN (só pro SISTEMA BUSCA NACIONAL).
//   • Tokens ficam nas Propriedades do Script e nos *secrets* do Streamlit
//     Cloud — NUNCA no código (os repositórios do GitHub são públicos).
//
// PUBLICAR (uma vez só, na conta dona da pasta CERTIFICADOS):
//   1. script.google.com > Novo projeto > colar este código > Salvar.
//   2. Selecionar a função `configurar` > Executar > autorizar o acesso ao
//      Drive. Ela acha a pasta CERTIFICADOS, cria as 6 subpastas e gera os
//      tokens. Ver o resultado em "Registro de execução".
//   3. Implantar > Nova implantação > "Aplicativo da Web".
//      Executar como: Eu. Quem pode acessar: Qualquer pessoa.
//   4. Em cada app do Streamlit Cloud (Settings > Secrets), colar:
//        CERT_DRIVE_URL   = "<URL da implantação>"
//        CERT_DRIVE_TOKEN = "<TOKEN_ daquele escritório>"
//
// Mudou o código? Colar aqui de novo (este arquivo é só backup) e
// "Implantar > Gerenciar implantações > lápis > Nova versão" pra manter a URL.
// Pra trocar um token vazado: apagar a propriedade TOKEN_<X> e rodar
// `configurar` de novo (gera só os que faltam).
// ============================================================================

const ESCRITORIOS = ["ASBEM", "VIDAL", "VS", "WR", "VISAO", "VJ"];
const NOME_PASTA_RAIZ = "CERTIFICADOS";
const NOME_SENHAS = "senhas.json";
const TAMANHO_MAX = 1024 * 1024;   // 1 MB — um .pfx tem poucos KB


// ── Rodar à mão pelo editor (passo 2) ────────────────────────────────────────
function configurar() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("PASTA_ID")) {
    const eu = Session.getEffectiveUser().getEmail();
    const it = DriveApp.getFoldersByName(NOME_PASTA_RAIZ);
    const achadas = [];
    while (it.hasNext()) {
      const f = it.next();
      const dono = f.getOwner();
      if (!f.isTrashed() && dono && dono.getEmail() === eu) achadas.push(f);
    }
    if (achadas.length !== 1) {
      throw new Error("Esperava 1 pasta '" + NOME_PASTA_RAIZ + "' sua no Drive, achei " +
                      achadas.length + ". Renomeie as outras ou grave o ID na propriedade PASTA_ID.");
    }
    props.setProperty("PASTA_ID", achadas[0].getId());
  }

  const raiz = pastaRaiz_();
  if (raiz.getSharingAccess() !== DriveApp.Access.PRIVATE) {
    Logger.log("⚠️ ATENÇÃO: a pasta CERTIFICADOS está compartilhada por link. Deixe-a restrita.");
  }
  ESCRITORIOS.forEach(function (esc) { subpasta_(esc); });

  const linhas = [];
  ESCRITORIOS.concat(["ADMIN"]).forEach(function (k) {
    const p = "TOKEN_" + k;
    if (!props.getProperty(p)) props.setProperty(p, Utilities.getUuid().replace(/-/g, ""));
    linhas.push(p + " = " + props.getProperty(p));
  });
  Logger.log("Pasta: " + raiz.getUrl() + "\nSubpastas: " + ESCRITORIOS.join(", ") +
             "\n\nTokens (copiar pros secrets de cada escritório):\n" + linhas.join("\n"));
}


// ── Web app ──────────────────────────────────────────────────────────────────
function doGet() {
  // Só pra conferir se a implantação está no ar — não devolve dado nenhum.
  return json_({ status: "ok", servico: "certificados-drive" });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const acao = req.acao || "upload";
    if (acao === "upload") return json_(upload_(req));
    if (acao === "listar") return json_(listar_(req));
    if (acao === "baixar") return json_(baixar_(req));
    return json_({ status: "erro", mensagem: "Ação desconhecida: " + acao });
  } catch (err) {
    return json_({ status: "erro", mensagem: String(err && err.message || err) });
  }
}


// ── upload: {token, escritorio, nome_arquivo, conteudo_b64, senha,
//             cnpj, razao_social, validade_iso} ─────────────────────────────
function upload_(req) {
  const esc = String(req.escritorio || "").toUpperCase();
  if (ESCRITORIOS.indexOf(esc) < 0) return { status: "erro", mensagem: "Escritório inválido: " + esc };
  if (!tokenOk_(req.token, "TOKEN_" + esc)) return { status: "erro", mensagem: "Token inválido." };

  const nome = String(req.nome_arquivo || "").replace(/[\\\/:*?"<>|]/g, "_").trim();
  if (!/\.(pfx|p12)$/i.test(nome)) return { status: "erro", mensagem: "Arquivo precisa ser .pfx ou .p12." };

  const bytes = Utilities.base64Decode(String(req.conteudo_b64 || ""));
  if (!bytes.length || bytes.length > TAMANHO_MAX) return { status: "erro", mensagem: "Arquivo vazio ou grande demais." };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);   // dois uploads do mesmo escritório mexem no mesmo senhas.json
  try {
    const pasta = subpasta_(esc);
    const md5 = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes));
    let resultado = "novo";

    const iguais = pasta.getFilesByName(nome);
    while (iguais.hasNext()) {
      const antigo = iguais.next();
      if (hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, antigo.getBlob().getBytes())) === md5) {
        resultado = "ja_existia";
      } else {
        antigo.setTrashed(true);   // mesmo nome, conteúdo diferente → fica o novo (vai pra lixeira, recuperável)
        resultado = "substituido";
      }
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

    const senhas = lerSenhas_(pasta);
    senhas[nome] = String(req.senha || "");
    gravarSenhas_(pasta, senhas);

    return { status: "ok", resultado: resultado, arquivo: nome, id: arq ? arq.getId() : "" };
  } finally {
    lock.releaseLock();
  }
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
  const id = PropertiesService.getScriptProperties().getProperty("PASTA_ID");
  if (!id) throw new Error("Rode a função configurar() no editor antes de usar.");
  return DriveApp.getFolderById(id);
}

function subpasta_(esc) {
  const raiz = pastaRaiz_();
  const it = raiz.getFoldersByName(esc);
  return it.hasNext() ? it.next() : raiz.createFolder(esc);
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
