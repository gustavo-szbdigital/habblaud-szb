import http from 'node:http';
import type { AddressInfo } from 'node:net';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { OfficeSnapshot } from '../../shared/types';
import { AccountsService } from '../accounts/service';
import { createApiHandler, type ApiDeps } from '../http/app';
import { createRequestGuard } from '../http/guard';
import { Hub } from '../http/sse';
import { setQuiet } from '../log';
import { NameStore } from '../model/names';
import { Office, type OfficeDeps } from '../model/office';
import { tempDir } from '../test/fixtures';
import { AGENTES_INICIAIS, AgentesRegistry } from './agentes';

setQuiet(true);

const TEST_TOKEN = 'secret-token-12345';

interface TestEnv {
  base: string;
  office: Office;
  agentes: AgentesRegistry;
  dataDir: string;
  close: () => Promise<void>;
  cleanup: () => void;
}

async function startServer(
  extra: Partial<ApiDeps> = {},
  officeExtra: Partial<OfficeDeps> = {},
  agentesToken?: string,
): Promise<TestEnv> {
  const tmp = tempDir();
  const late: { office?: Office } = {};
  const accounts = new AccountsService({ dirs: [], home: tmp.dir, env: {}, onChange: () => late.office?.markDirty() });
  const office = new Office({
    names: new NameStore(null),
    version: '1.0.0',
    startedAt: Date.now(),
    accounts: (s) => accounts.list(s),
    sources: () => [],
    accountName: () => undefined,
    ...officeExtra,
  });
  late.office = office;
  const hub = new Hub(office, { throttleMs: 10 });
  hub.start();

  const agentes = new AgentesRegistry({ dataDir: tmp.dir, office });
  agentes.load();
  agentes.start();

  const api = createApiHandler({
    office,
    hub,
    accounts,
    sources: () => agentes.sources(),
    version: '1.0.0',
    inDocker: false,
    agentes,
    agentesToken,
    ...extra,
  });

  const guard = createRequestGuard({ allowedHosts: new Set() });
  const server = http.createServer((req, res) => {
    if (guard(req, res)) return;
    const url = new URL(req.url ?? '/', 'http://127.0.0.1');
    if (!api(req, res, url)) {
      res.writeHead(404);
      res.end();
    }
  });

  await new Promise<void>((ok) => server.listen(0, '127.0.0.1', ok));
  const port = (server.address() as AddressInfo).port;

  return {
    base: `http://127.0.0.1:${port}`,
    office,
    agentes,
    dataDir: tmp.dir,
    close: () =>
      new Promise((ok) => {
        hub.stop();
        server.closeAllConnections();
        server.close(() => ok());
      }),
    cleanup: tmp.cleanup,
  };
}

async function snapshotOf(base: string): Promise<OfficeSnapshot> {
  const res = await fetch(`${base}/api/snapshot`);
  return (await res.json()) as OfficeSnapshot;
}

function requestRaw(
  base: string,
  path: string,
  opts: { method?: string; headers?: Record<string, string>; body?: string } = {},
): Promise<{ status: number; body: string; headers: http.IncomingHttpHeaders }> {
  const u = new URL(base);
  return new Promise((ok, fail) => {
    const req = http.request(
      {
        host: u.hostname,
        port: u.port,
        path,
        method: opts.method ?? 'GET',
        headers: opts.headers,
      },
      (res) => {
        let body = '';
        res.setEncoding('utf8');
        res.on('data', (c: string) => (body += c));
        res.on('end', () => ok({ status: res.statusCode ?? 0, body, headers: res.headers }));
      },
    );
    req.on('error', fail);
    req.end(opts.body);
  });
}

describe('Registro de Agentes Fixos (server/sources/agentes.ts)', () => {
  it('cria agentes iniciais com status idle e persiste em JSON em HABBLAUD_DATA_DIR', () => {
    const tmp = tempDir();
    try {
      const filePath = join(tmp.dir, 'agentes.json');
      expect(existsSync(filePath)).toBe(false);

      const registry = new AgentesRegistry({ dataDir: tmp.dir });
      registry.load();

      expect(existsSync(filePath)).toBe(true);
      const list = registry.list();
      expect(list.length).toBe(4);

      const ids = list.map((a) => a.id);
      expect(ids).toEqual(['gestor-trafego', 'redator', 'designer', 'seo']);

      for (const a of list) {
        expect(a.status).toBe('idle');
        expect(a.sala).toBe('marketing');
      }

      // Sobrevive a reinício
      registry.applyEvent({
        agente: 'gestor-trafego',
        status: 'working',
        atividade: 'Otimizando campanhas',
      });

      const reloaded = new AgentesRegistry({ dataDir: tmp.dir });
      reloaded.load();

      const gt = reloaded.get('gestor-trafego');
      expect(gt?.status).toBe('working');
      expect(gt?.activity?.text).toBe('Otimizando campanhas');
      expect(gt?.recent?.length).toBe(1);
    } finally {
      tmp.cleanup();
    }
  });

  it('preserva dados persistidos válidos se arquivo já existir', () => {
    const tmp = tempDir();
    try {
      const filePath = join(tmp.dir, 'agentes.json');
      const customAgents = [
        {
          id: 'gestor-trafego',
          nome: 'Gestor de Tráfego',
          papel: 'Gestor de Tráfego',
          sala: 'marketing',
          status: 'working',
          waitingFor: undefined,
        },
      ];
      const { writeFileSync } = require('node:fs');
      writeFileSync(filePath, JSON.stringify(customAgents, null, 2), 'utf8');

      const registry = new AgentesRegistry({ dataDir: tmp.dir });
      registry.load();

      expect(registry.get('gestor-trafego')?.status).toBe('working');
      // Garante que os outros agentes iniciais que faltavam são completados
      expect(registry.get('redator')?.status).toBe('idle');
      expect(registry.get('designer')?.status).toBe('idle');
      expect(registry.get('seo')?.status).toBe('idle');
    } finally {
      tmp.cleanup();
    }
  });
});

describe('Endpoint POST /api/agentes/evento e Integração com Office', () => {
  let env: TestEnv;

  beforeEach(async () => {
    env = await startServer({}, {}, TEST_TOKEN);
  });

  afterEach(async () => {
    await env.close();
    env.cleanup();
  });

  it('agentes aparecem no snapshot (/api/snapshot) com status inicial idle mesmo sem Claude Code', async () => {
    const snap = await snapshotOf(env.base);

    expect(snap.rooms.some((r) => r.id === 'marketing')).toBe(true);

    const gestor = snap.agents.find((a) => a.id === 'gestor-trafego');
    expect(gestor).toBeDefined();
    expect(gestor?.name).toBe('Gestor de Tráfego');
    expect(gestor?.role).toBe('Gestor de Tráfego');
    expect(gestor?.roomId).toBe('marketing');
    expect(gestor?.status).toBe('idle');

    const redator = snap.agents.find((a) => a.id === 'redator');
    expect(redator).toBeDefined();
    expect(redator?.name).toBe('Redator');
    expect(redator?.status).toBe('idle');

    const designer = snap.agents.find((a) => a.id === 'designer');
    expect(designer).toBeDefined();
    expect(designer?.name).toBe('Designer');
    expect(designer?.status).toBe('idle');

    const seo = snap.agents.find((a) => a.id === 'seo');
    expect(seo).toBeDefined();
    expect(seo?.name).toBe('SEO');
    expect(seo?.status).toBe('idle');
  });

  it('evento válido atualiza status para working e publica atividade em recent', async () => {
    const res = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${TEST_TOKEN}`,
      },
      body: JSON.stringify({
        agente: 'gestor-trafego',
        status: 'working',
        atividade: 'Subindo anúncios no Meta Ads',
      }),
    });

    expect(res.status).toBe(200);
    expect(JSON.parse(res.body)).toEqual({ ok: true });

    const snap = await snapshotOf(env.base);
    const gestor = snap.agents.find((a) => a.id === 'gestor-trafego');
    expect(gestor?.status).toBe('working');
    expect(gestor?.activity?.text).toBe('Subindo anúncios no Meta Ads');
    expect(gestor?.recent.length).toBeGreaterThanOrEqual(1);
    expect(gestor?.recent[gestor.recent.length - 1].text).toBe('Subindo anúncios no Meta Ads');
  });

  it('evento com status waiting atualiza status e esperandoPor', async () => {
    const res = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${TEST_TOKEN}`,
      },
      body: JSON.stringify({
        agente: 'designer',
        status: 'waiting',
        atividade: 'Aguardando aprovação do layout',
        esperandoPor: 'aprovação do cliente',
      }),
    });

    expect(res.status).toBe(200);
    expect(JSON.parse(res.body)).toEqual({ ok: true });

    const snap = await snapshotOf(env.base);
    const designer = snap.agents.find((a) => a.id === 'designer');
    expect(designer?.status).toBe('waiting');
    expect(designer?.waitingFor).toBe('aprovação do cliente');
    expect(designer?.activity?.text).toBe('Aguardando aprovação do layout');
  });

  it('retorna 400 para JSON inválido, corpo malformado ou campos incorretos', async () => {
    // JSON malformado
    const r1 = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${TEST_TOKEN}`,
      },
      body: '{ json invalido',
    });
    expect(r1.status).toBe(400);

    // Sem campo agente
    const r2 = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${TEST_TOKEN}`,
      },
      body: JSON.stringify({
        status: 'working',
        atividade: 'Fazendo algo',
      }),
    });
    expect(r2.status).toBe(400);

    // Status inválido
    const r3 = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${TEST_TOKEN}`,
      },
      body: JSON.stringify({
        agente: 'gestor-trafego',
        status: 'dormindo',
        atividade: 'Fazendo algo',
      }),
    });
    expect(r3.status).toBe(400);

    // Atividade vazia
    const r4 = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${TEST_TOKEN}`,
      },
      body: JSON.stringify({
        agente: 'gestor-trafego',
        status: 'working',
        atividade: '   ',
      }),
    });
    expect(r4.status).toBe(400);
  });

  it('retorna 400 para agente desconhecido', async () => {
    const res = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${TEST_TOKEN}`,
      },
      body: JSON.stringify({
        agente: 'agente-fantasma',
        status: 'working',
        atividade: 'Assombrando o escritório',
      }),
    });

    expect(res.status).toBe(400);
    const body = JSON.parse(res.body);
    expect(body.error).toContain('agente-fantasma');
  });

  it('retorna 401 quando Authorization está ausente ou inválido', async () => {
    // Sem Authorization
    const r1 = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        agente: 'gestor-trafego',
        status: 'working',
        atividade: 'Trabalhando',
      }),
    });
    expect(r1.status).toBe(401);

    // Token errado
    const r2 = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer token-errado',
      },
      body: JSON.stringify({
        agente: 'gestor-trafego',
        status: 'working',
        atividade: 'Trabalhando',
      }),
    });
    expect(r2.status).toBe(401);

    // Sem prefixo Bearer
    const r3 = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: TEST_TOKEN,
      },
      body: JSON.stringify({
        agente: 'gestor-trafego',
        status: 'working',
        atividade: 'Trabalhando',
      }),
    });
    expect(r3.status).toBe(401);
  });

  it('retorna 405 para métodos diferentes de POST', async () => {
    const rGet = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'GET',
      headers: { Authorization: `Bearer ${TEST_TOKEN}` },
    });
    expect(rGet.status).toBe(405);
    expect(rGet.headers.allow).toBe('POST');

    const rPut = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${TEST_TOKEN}` },
    });
    expect(rPut.status).toBe(405);
    expect(rPut.headers.allow).toBe('POST');

    const rDelete = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${TEST_TOKEN}` },
    });
    expect(rDelete.status).toBe(405);
    expect(rDelete.headers.allow).toBe('POST');
  });
});

describe('Endpoint POST /api/agentes/evento com token desativado (503)', () => {
  let env: TestEnv;

  beforeEach(async () => {
    // Inicia sem nenhum token definido (agentesToken = undefined)
    env = await startServer({}, {}, undefined);
  });

  afterEach(async () => {
    await env.close();
    env.cleanup();
  });

  it('responde 503 quando HABBLAUD_AGENTES_TOKEN não estiver definida', async () => {
    const res = await requestRaw(env.base, '/api/agentes/evento', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer qualquer-coisa',
      },
      body: JSON.stringify({
        agente: 'gestor-trafego',
        status: 'working',
        atividade: 'Tentando atualizar',
      }),
    });

    expect(res.status).toBe(503);
    const body = JSON.parse(res.body);
    expect(body.error).toContain('HABBLAUD_AGENTES_TOKEN');
  });
});
