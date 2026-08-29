// Edge Function: reset-technician-password
//
// Gera uma nova senha aleatória para um técnico e a grava no Supabase Auth.
// Existe porque o navegador (chave anon/publishable) NUNCA pode ter acesso
// à service role key — só ela consegue trocar a senha de outro usuário via
// supabase.auth.admin.updateUserById. Esta função roda no servidor do
// Supabase, guarda a service role key só como variável de ambiente (nunca
// enviada ao cliente), e confere que quem está chamando é realmente um
// admin (lendo o token de quem chamou, não confiando em nada vindo do body)
// antes de trocar a senha de outra pessoa.
//
// Deploy: supabase functions deploy reset-technician-password
// Chamada pelo site via: supabaseClient.functions.invoke('reset-technician-password', { body: { technicianId } })

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

const SUPABASE_URL = Deno.env.get('SUPABASE_URL')!;
const SUPABASE_ANON_KEY = Deno.env.get('SUPABASE_ANON_KEY')!;
const SERVICE_ROLE_KEY = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
};

function gerarSenha(): string {
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  let senha = '';
  for (let i = 0; i < 10; i++) {
    senha += chars[Math.floor(Math.random() * chars.length)];
  }
  return senha;
}

Deno.serve(async (req: Request) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: CORS_HEADERS });
  }

  try {
    const authHeader = req.headers.get('Authorization');
    if (!authHeader) {
      throw new Error('Requisição sem autenticação.');
    }

    // Cliente com o token de quem chamou — só serve pra descobrir QUEM é.
    const callerClient = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
      global: { headers: { Authorization: authHeader } },
    });
    const { data: { user: caller }, error: callerErr } = await callerClient.auth.getUser();
    if (callerErr || !caller) {
      throw new Error('Sessão inválida ou expirada.');
    }

    // Cliente com a service role key — ignora RLS, só usado depois de
    // confirmarmos que quem chamou é admin.
    const adminClient = createClient(SUPABASE_URL, SERVICE_ROLE_KEY);

    const { data: callerProfile, error: profileErr } = await adminClient
      .from('profiles')
      .select('role')
      .eq('id', caller.id)
      .single();

    if (profileErr || !callerProfile || callerProfile.role !== 'admin') {
      throw new Error('Só administradores podem redefinir a senha de outro usuário.');
    }

    const body = await req.json().catch(() => ({}));
    const technicianId = body?.technicianId;
    if (!technicianId || typeof technicianId !== 'string') {
      throw new Error('technicianId é obrigatório.');
    }

    const novaSenha = gerarSenha();
    const { error: updateErr } = await adminClient.auth.admin.updateUserById(technicianId, {
      password: novaSenha,
    });
    if (updateErr) {
      throw updateErr;
    }

    return new Response(JSON.stringify({ senha: novaSenha }), {
      status: 200,
      headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return new Response(JSON.stringify({ error: message }), {
      status: 400,
      headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' },
    });
  }
});
