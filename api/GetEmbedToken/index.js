const https = require('https');
const querystring = require('querystring');

const TENANT_ID     = process.env.TENANT_ID     || '9e7ad80a-3510-4291-8108-f4da12e8e35d';
const CLIENT_ID     = process.env.CLIENT_ID     || 'f59c96e3-b802-4cf4-9bd4-3a9817b55e05';
const CLIENT_SECRET = process.env.CLIENT_SECRET || '41d9b450-6cf2-4627-82da-a6bfb8ca699a';
const WORKSPACE_ID  = process.env.WORKSPACE_ID  || '109eff3f-c15c-4a00-be45-758b4ceb4ecd';
const REPORT_ID     = process.env.REPORT_ID     || '84e44d37-9922-4a17-aaae-30bea1288a0b';

function httpsPost(hostname, path, headers, body) {
  return new Promise((resolve, reject) => {
    const req = https.request({ hostname, path, method: 'POST', headers }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(data) }); }
        catch (e) { resolve({ status: res.statusCode, body: data }); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

module.exports = async function (context, req) {
  context.res = { headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' } };

  try {
    // Step 1: Get AAD token
    const tokenBody = querystring.stringify({
      grant_type:    'client_credentials',
      client_id:     CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope:         'https://analysis.windows.net/powerbi/api/.default',
    });

    const tokenRes = await httpsPost(
      'login.microsoftonline.com',
      `/${TENANT_ID}/oauth2/v2.0/token`,
      { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(tokenBody) },
      tokenBody
    );

    if (tokenRes.status !== 200) {
      context.res.status = 502;
      context.res.body = JSON.stringify({ error: 'AAD token failed', detail: tokenRes.body });
      return;
    }

    const accessToken = tokenRes.body.access_token;

    // Step 2: Get Power BI embed token
    const embedBody = JSON.stringify({ accessLevel: 'View' });
    const embedRes = await httpsPost(
      'api.powerbi.com',
      `/v1.0/myorg/groups/${WORKSPACE_ID}/reports/${REPORT_ID}/GenerateToken`,
      {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type':  'application/json',
        'Content-Length': Buffer.byteLength(embedBody),
      },
      embedBody
    );

    if (embedRes.status !== 200) {
      context.res.status = 502;
      context.res.body = JSON.stringify({ error: 'Embed token failed', detail: embedRes.body });
      return;
    }

    context.res.status = 200;
    context.res.body = JSON.stringify({
      embedToken:  embedRes.body.token,
      tokenExpiry: embedRes.body.expiration,
      reportId:    REPORT_ID,
      workspaceId: WORKSPACE_ID,
      embedUrl:    `https://app.powerbi.com/reportEmbed?reportId=${REPORT_ID}&groupId=${WORKSPACE_ID}`,
    });

  } catch (err) {
    context.res.status = 500;
    context.res.body = JSON.stringify({ error: err.message });
  }
};
