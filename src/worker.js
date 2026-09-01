export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === '/api/fetch-pricelist' && request.method === 'POST') {
      return handleFetchPricelist(request);
    }

    return env.ASSETS.fetch(request);
  }
};

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
        error: 'Fournisseur non configuré pour la récupération en ligne',
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
              return new Response(JSON.stringify({ fichier: fichier.nom, url: fichier.url, type: 'binary', message: 'Fichier binaire — téléchargement direct requis' }), { headers });
            }
          }
        }
      }

      return new Response(JSON.stringify({ fichiers, base: config.base }), { headers });
    }

    return new Response(JSON.stringify({ error: 'Format non supporté' }), { status: 500, headers });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), { status: 500, headers });
  }
}
