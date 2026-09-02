export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === '/api/fetch-pricelist' && request.method === 'POST') {
      return handleFetchPricelist(request);
    }

    if (url.pathname === '/api/detect-columns' && request.method === 'POST') {
      return handleDetectColumns(request, env);
    }

    return env.ASSETS.fetch(request);
  }
};

async function handleDetectColumns(request, env) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  };

  try {
    if (!env.CLAUDE_API_KEY) {
      return new Response(JSON.stringify({ error: 'CLAUDE_API_KEY non configuree' }), { status: 500, headers });
    }

    const body = await request.json();
    const { lignes } = body;

    if (!lignes || !lignes.length) {
      return new Response(JSON.stringify({ error: 'Aucune ligne fournie' }), { status: 400, headers });
    }

    const echantillon = lignes.slice(0, 20).join('\n');

    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': env.CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 300,
        messages: [{
          role: 'user',
          content: `Voici les premieres lignes d'une liste de prix fournisseur (separees par tabulation ou virgule). Identifie les index de colonnes (base 0) pour:
- code: le code article, numero de piece, part number, SKU, catalogue number
- prix: le prix unitaire, prix coutant, list price
- desc: la description du produit

Reponds UNIQUEMENT en JSON strict, rien d'autre:
{"code": INDEX, "prix": INDEX, "desc": INDEX, "entete": INDEX_LIGNE}

ou INDEX est le numero de colonne (base 0), et entete est le numero de la ligne d'en-tete (base 0, -1 si pas d'en-tete).
Si une colonne est introuvable, mets -1.

Donnees:
${echantillon}`
        }]
      })
    });

    if (!resp.ok) {
      const errText = await resp.text();
      return new Response(JSON.stringify({ error: 'Erreur Claude API: ' + resp.status, details: errText }), { status: 502, headers });
    }

    const data = await resp.json();
    const texte = data.content[0].text.trim();

    const jsonMatch = texte.match(/\{[^}]+\}/);
    if (!jsonMatch) {
      return new Response(JSON.stringify({ error: 'Reponse Claude invalide', raw: texte }), { status: 500, headers });
    }

    const colonnes = JSON.parse(jsonMatch[0]);
    return new Response(JSON.stringify(colonnes), { headers });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), { status: 500, headers });
  }
}

const FOURNISSEURS_CONNUS = {
  'cb supplies': {
    base: 'https://cbsupplies.ca/sheets/en/',
    format: 'html-links'
  }
};

async function handleFetchPricelist(request) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  };

  try {
    const body = await request.json();
    const { fournisseur, categorie } = body;

    if (!fournisseur) {
      return new Response(JSON.stringify({ error: 'Fournisseur requis' }), { status: 400, headers });
    }

    const fournLower = fournisseur.toLowerCase().trim();
    let config = null;
    for (const [key, val] of Object.entries(FOURNISSEURS_CONNUS)) {
      if (fournLower.includes(key)) { config = val; break; }
    }

    if (!config) {
      return new Response(JSON.stringify({
        error: 'Fournisseur non configure pour la recuperation en ligne',
        fournisseur,
        connus: Object.keys(FOURNISSEURS_CONNUS)
      }), { status: 404, headers });
    }

    if (config.format === 'html-links') {
      const indexResp = await fetch(config.base, {
        headers: { 'User-Agent': 'GromecPriceBot/1.0' }
      });
      if (!indexResp.ok) {
        return new Response(JSON.stringify({ error: 'Impossible de joindre ' + config.base, status: indexResp.status }), { status: 502, headers });
      }

      const html = await indexResp.text();
      const linkRegex = /href=["']([^"']+\.(?:csv|txt|xlsx?|pdf))["']/gi;
      const fichiers = [];
      let m;
      while ((m = linkRegex.exec(html)) !== null) {
        const href = m[1];
        const fullUrl = href.startsWith('http') ? href : new URL(href, config.base).href;
        const nom = decodeURIComponent(href.split('/').pop());
        fichiers.push({ nom, url: fullUrl });
      }

      if (categorie) {
        const catNorm = categorie.toUpperCase().replace(/[\s\-]/g, '');
        const filtres = fichiers.filter(f => {
          const nomNorm = f.nom.toUpperCase().replace(/[\s\-]/g, '');
          return nomNorm.includes(catNorm);
        });
        if (filtres.length > 0) {
          const fichier = filtres[0];
          const dataResp = await fetch(fichier.url, {
            headers: { 'User-Agent': 'GromecPriceBot/1.0' }
          });
          if (dataResp.ok) {
            const contentType = dataResp.headers.get('content-type') || '';
            if (contentType.includes('text') || fichier.nom.endsWith('.csv') || fichier.nom.endsWith('.txt')) {
              const texte = await dataResp.text();
              return new Response(JSON.stringify({ fichier: fichier.nom, contenu: texte, type: 'text' }), { headers });
            } else {
              return new Response(JSON.stringify({ fichier: fichier.nom, url: fichier.url, type: 'binary', message: 'Fichier binaire — telechargement direct requis' }), { headers });
            }
          }
        }
      }

      return new Response(JSON.stringify({ fichiers, base: config.base }), { headers });
    }

    return new Response(JSON.stringify({ error: 'Format non supporte' }), { status: 500, headers });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), { status: 500, headers });
  }
}
