// DEBBIOM Edge Function: create-user
// Crea usuarios con @debbiom.app usando service_role (server-side)
// Solo admins autenticados pueden llamar esta funcion

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { serve } from 'https://deno.land/std@0.177.0/http/server.ts'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

serve(async (req: Request) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const serviceRoleKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
    const anonKey = Deno.env.get('SUPABASE_ANON_KEY')!

    // 1. Verificar que el llamador es admin
    const authHeader = req.headers.get('Authorization')
    if (!authHeader) {
      return new Response(JSON.stringify({ error: 'No autorizado' }), {
        status: 401, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    // Cliente con la sesion del llamador para verificar identidad
    const callerClient = createClient(supabaseUrl, anonKey, {
      global: { headers: { Authorization: authHeader } }
    })

    const { data: { user: caller }, error: authError } = await callerClient.auth.getUser()
    if (authError || !caller) {
      return new Response(JSON.stringify({ error: 'Sesion invalida' }), {
        status: 401, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    // Verificar rol admin en profiles
    const { data: profile } = await callerClient.from('profiles').select('role').eq('id', caller.id).single()
    if (!profile || profile.role !== 'admin') {
      return new Response(JSON.stringify({ error: 'Solo administradores pueden crear usuarios' }), {
        status: 403, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    // 2. Leer datos del nuevo usuario
    const { username, password, fullName, role } = await req.json()

    if (!username || !password) {
      return new Response(JSON.stringify({ error: 'Faltan datos obligatorios' }), {
        status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    // Validaciones
    if (!/^[a-zA-Z0-9._-]{2,30}$/.test(username.trim())) {
      return new Response(JSON.stringify({ error: 'Usuario invalido (2-30 chars, alfanumerico)' }), {
        status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    if (password.length < 8 || password.length > 128) {
      return new Response(JSON.stringify({ error: 'Contrasena debe tener 8-128 caracteres' }), {
        status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    const safeRole = ['admin', 'user'].includes(role) ? role : 'user'
    const email = username.trim().toLowerCase() + '@debbiom.app'

    // 3. Crear usuario con service_role (bypasses email restrictions)
    const adminClient = createClient(supabaseUrl, serviceRoleKey)

    const { data: newUser, error: createError } = await adminClient.auth.admin.createUser({
      email: email,
      password: password,
      email_confirm: true,
      user_metadata: { username: username.trim(), role: safeRole, full_name: fullName || username }
    })

    if (createError) {
      return new Response(JSON.stringify({ error: 'Error al crear: ' + createError.message }), {
        status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    // 4. Crear perfil
    if (newUser.user) {
      const sanitize = (s: string) => s.replace(/[<>"'&]/g, (c: string) =>
        ({ '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;', '&': '&amp;' }[c] || c))

      await adminClient.from('profiles').insert({
        id: newUser.user.id,
        username: sanitize(username.trim()),
        role: safeRole,
        full_name: sanitize(fullName || username)
      })
    }

    return new Response(JSON.stringify({
      success: true,
      userId: newUser.user?.id
    }), {
      status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    })

  } catch (err) {
    return new Response(JSON.stringify({ error: 'Error interno del servidor' }), {
      status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    })
  }
})
