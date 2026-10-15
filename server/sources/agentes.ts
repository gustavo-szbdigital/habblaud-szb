// Registro de agentes fixos (fonte n8n): persistido em <dataDir>/agentes.json para sobreviver
// a reinícios do servidor, alimentando o Office e recebendo eventos via POST /api/agentes/evento.
import { mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import type { Activity, ActivityKind, AgentStatus, SourceInfo } from '../../shared/types';
import type { Office } from '../model/office';
import { errMsg, log } from '../log';
import { tr } from '../../shared/i18n';

export interface AgenteFixo {
  id: string;
  nome: string;
  papel: string;
  sala: string;
  status: AgentStatus;
  waitingFor?: string;
  activity?: Activity;
  recent?: Activity[];
  look?: 'f' | 'm';
}

export interface AgenteEvento {
  agente: string;
  status: 'working' | 'idle' | 'waiting';
  atividade: string;
  esperandoPor?: string;
}

export interface AgentesOptions {
  /** Diretório de dados (HABBLAUD_DATA_DIR ou ~/.habblaud). Se omitido, usa process.env ou default. */
  dataDir?: string;
  /** Caminho exato do arquivo JSON. Se for null, roda apenas em memória (útil para testes). */
  file?: string | null;
  /** Instância do Office para alimentar. */
  office?: Office;
  now?: () => number;
}

export const AGENTES_INICIAIS: readonly AgenteFixo[] = [
  {
    id: 'gestor-trafego',
    nome: 'Gestor de Tráfego',
    papel: 'Gestor de Tráfego',
    sala: 'marketing',
    status: 'idle',
    look: 'm',
  },
  {
    id: 'redator',
    nome: 'Redator',
    papel: 'Redator',
    sala: 'marketing',
    status: 'idle',
    look: 'f',
  },
  {
    id: 'designer',
    nome: 'Designer',
    papel: 'Designer',
    sala: 'marketing',
    status: 'idle',
    look: 'm',
  },
  {
    id: 'seo',
    nome: 'SEO',
    papel: 'SEO',
    sala: 'marketing',
    status: 'idle',
    look: 'f',
  },
];

function cloneInitialAgents(): AgenteFixo[] {
  return AGENTES_INICIAIS.map((a) => ({ ...a, recent: [] }));
}

export class AgentesRegistry {
  private agentes = new Map<string, AgenteFixo>();
  private readonly file: string | null;
  private readonly office?: Office;
  private readonly now: () => number;
  private seq = 0;

  constructor(opts: AgentesOptions = {}) {
    this.office = opts.office;
    this.now = opts.now ?? Date.now;

    if (opts.file !== undefined) {
      this.file = opts.file;
    } else {
      const dataDir = resolve(opts.dataDir || process.env.HABBLAUD_DATA_DIR?.trim() || join(homedir(), '.habblaud'));
      this.file = join(dataDir, 'agentes.json');
    }
  }

  load(): void {
    if (!this.file) {
      for (const a of cloneInitialAgents()) {
        this.agentes.set(a.id, a);
      }
      return;
    }

    try {
      const raw = readFileSync(this.file, 'utf8');
      const parsed = JSON.parse(raw) as unknown;
      const list = Array.isArray(parsed)
        ? parsed
        : parsed && typeof parsed === 'object' && Array.isArray((parsed as { agentes?: unknown }).agentes)
          ? (parsed as { agentes: unknown[] }).agentes
          : [];

      if (!list.length) {
        for (const a of cloneInitialAgents()) {
          this.agentes.set(a.id, a);
        }
        this.saveSync();
        return;
      }

      for (const item of list) {
        if (!item || typeof item !== 'object') continue;
        const entry = item as Record<string, unknown>;
        const id = typeof entry.id === 'string' ? entry.id.trim() : '';
        if (!id) continue;

        const nome = typeof entry.nome === 'string' ? entry.nome : typeof entry.name === 'string' ? entry.name : id;
        const papel = typeof entry.papel === 'string' ? entry.papel : typeof entry.role === 'string' ? entry.role : nome;
        const sala = typeof entry.sala === 'string' ? entry.sala : typeof entry.room === 'string' ? entry.room : typeof entry.roomId === 'string' ? entry.roomId : 'marketing';
        const rawStatus = typeof entry.status === 'string' ? entry.status : 'idle';
        const status: AgentStatus = rawStatus === 'working' || rawStatus === 'waiting' || rawStatus === 'idle' ? rawStatus : 'idle';
        const waitingFor = typeof entry.waitingFor === 'string' ? entry.waitingFor : typeof entry.esperandoPor === 'string' ? entry.esperandoPor : undefined;
        const look = entry.look === 'f' || entry.look === 'm' ? entry.look : undefined;
        const recent = Array.isArray(entry.recent) ? (entry.recent as Activity[]) : [];
        const activity = entry.activity && typeof entry.activity === 'object' ? (entry.activity as Activity) : recent[recent.length - 1];

        this.agentes.set(id, {
          id,
          nome,
          papel,
          sala,
          status,
          waitingFor,
          look,
          recent,
          activity,
        });
      }

      // Se faltar algum agente inicial essencial, garante que existe
      for (const init of AGENTES_INICIAIS) {
        if (!this.agentes.has(init.id)) {
          this.agentes.set(init.id, { ...init, recent: [] });
        }
      }
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code === 'ENOENT') {
        for (const a of cloneInitialAgents()) {
          this.agentes.set(a.id, a);
        }
        this.saveSync();
      } else {
        log.warn(tr('agentes.json ilegível ({0}); usando agentes padrão.', [errMsg(err)]));
        for (const a of cloneInitialAgents()) {
          this.agentes.set(a.id, a);
        }
      }
    }
  }

  saveSync(): void {
    if (!this.file) return;
    try {
      const payload = this.list();
      const tmp = `${this.file}.${process.pid}.tmp`;
      const data = JSON.stringify(payload, null, 2);
      mkdirSync(dirname(this.file), { recursive: true });
      writeFileSync(tmp, data, 'utf8');
      renameSync(tmp, this.file);
    } catch (err) {
      log.warn(tr('Falha ao persistir agentes.json: {0}', [errMsg(err)]));
    }
  }

  start(): void {
    if (!this.office) return;
    const now = this.now();
    for (const a of this.agentes.values()) {
      const cwd = a.sala;
      this.office.addMain({
        id: a.id,
        name: a.nome,
        look: a.look ?? 'm',
        role: a.papel,
        account: 'n8n',
        sessionId: a.id,
        cwd,
        status: a.status,
        waitingFor: a.waitingFor,
        startedAt: now,
      });

      if (a.recent && a.recent.length > 0) {
        for (const act of a.recent) {
          this.office.addActivity(a.id, act, act.id === a.activity?.id, { feed: false });
        }
      } else if (a.activity) {
        this.office.addActivity(a.id, a.activity, true, { feed: false });
      }
    }
  }

  sources(): SourceInfo[] {
    return [
      {
        label: 'n8n',
        path: this.file ?? 'agentes.json',
        sessions: this.agentes.size,
        ok: true,
      },
    ];
  }

  get(id: string): AgenteFixo | undefined {
    return this.agentes.get(id);
  }

  list(): AgenteFixo[] {
    return [...this.agentes.values()];
  }

  applyEvent(ev: AgenteEvento): AgenteFixo {
    const agente = this.agentes.get(ev.agente);
    if (!agente) {
      throw new Error(tr('agente desconhecido: {0}', [ev.agente]));
    }

    let kind: ActivityKind = 'other';
    let icon = '⚡';
    if (ev.status === 'working') {
      kind = 'run';
      icon = '⚡';
    } else if (ev.status === 'waiting') {
      kind = 'wait';
      icon = '✋';
    } else if (ev.status === 'idle') {
      kind = 'done';
      icon = '✅';
    }

    const now = this.now();
    const actId = `${agente.id}#${now}-${++this.seq}`;
    const activity: Activity = {
      id: actId,
      kind,
      icon,
      text: ev.atividade,
      at: now,
    };
    if (ev.status === 'waiting' && ev.esperandoPor) {
      activity.detail = ev.esperandoPor;
    }

    agente.status = ev.status;
    if (ev.status === 'waiting') {
      agente.waitingFor = ev.esperandoPor || tr('aguardando');
    } else {
      delete agente.waitingFor;
    }
    agente.activity = activity;
    agente.recent = [...(agente.recent ?? []), activity].slice(-30);

    if (this.office) {
      this.office.setStatus(agente.id, ev.status, agente.waitingFor);
      this.office.addActivity(agente.id, activity, true, { feed: true });
    }

    this.saveSync();
    return agente;
  }
}
