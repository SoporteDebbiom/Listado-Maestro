import { createClient } from 'https://esm.sh/@supabase/supabase-js@2.49.1'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
}

Deno.serve(async (req: Request) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const supabaseAnonKey = Deno.env.get('SUPABASE_ANON_KEY')!
    const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!

    // SEGURIDAD: Solo autenticación JWT. Bootstrap key eliminado (CRIT-04 fix).
    const authHeader = req.headers.get('Authorization')

    let adminUserId = ''
    let adminUsername = 'unknown'

    if (!authHeader) {
      return new Response(JSON.stringify({ error: 'No autorizado. Se requiere sesión activa.' }), {
        status: 401, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    // JWT mode: authenticated via user session
    const userClient = createClient(supabaseUrl, supabaseAnonKey, {
      global: { headers: { Authorization: authHeader } },
      auth: { persistSession: false }
    })

    const { data: { user }, error: authError } = await userClient.auth.getUser()
    if (authError || !user) {
      return new Response(JSON.stringify({ error: 'Sesion invalida' }), {
        status: 401, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    const { data: profile, error: profileError } = await userClient
      .from('profiles')
      .select('role, username')
      .eq('id', user.id)
      .single()

    if (profileError || !profile || profile.role !== 'admin') {
      return new Response(JSON.stringify({ error: 'Sin permisos de administrador' }), {
        status: 403, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }
    adminUserId = user.id
    adminUsername = profile.username || 'unknown'

    const adminClient = createClient(supabaseUrl, supabaseServiceKey, {
      auth: { persistSession: false }
    })

    const body = await req.json()
    const { action } = body

    async function _audit(act: string, detail: string) {
      try {
        await adminClient.from('audit_log').insert({
          admin_id: adminUserId,
          admin_username: adminUsername,
          action: act,
          detail: detail,
          ip: req.headers.get('x-forwarded-for') || req.headers.get('x-real-ip') || 'unknown',
          created_at: new Date().toISOString()
        })
      } catch (_) {}
    }

    switch (action) {
      case 'createUser': {
        const { username, password, fullName, role } = body
        if (!username || !password) {
          return new Response(JSON.stringify({ error: 'Datos incompletos' }), {
            status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        if (typeof username !== 'string' || !/^[a-zA-Z0-9._-]{2,30}$/.test(username.trim())) {
          return new Response(JSON.stringify({ error: 'Usuario invalido' }), {
            status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        if (typeof password !== 'string' || password.length < 8 || password.length > 128 ||
            !/[A-Z]/.test(password) || !/[a-z]/.test(password) || !/[0-9]/.test(password) ||
            !/[!@#$%^&*()_+\-=\[\]{};':"\\|,.<>\/?]/.test(password)) {
          return new Response(JSON.stringify({ error: 'Contrasena: min 8, mayuscula, minuscula, numero y caracter especial' }), {
            status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        const safeRole = (role === 'admin') ? 'admin' : 'user'
        const email = username.trim().toLowerCase() + '@debbiom.app'
        const safeName = (fullName || username).replace(/[<>"'&]/g, '')

        const { data: newUser, error: createError } = await adminClient.auth.admin.createUser({
          email,
          password,
          email_confirm: true,
          user_metadata: { username: username.trim(), role: safeRole, full_name: safeName }
        })
        if (createError) {
          return new Response(JSON.stringify({ error: 'Error al crear usuario: ' + createError.message }), {
            status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        const uid = newUser.user?.id
        if (uid) {
          const { error: profileError } = await adminClient.from('profiles').insert({
            id: uid, username: username.trim(), role: safeRole, full_name: safeName
          })
          if (profileError) {
            return new Response(JSON.stringify({ error: 'Usuario creado pero error en perfil: ' + profileError.message }), {
              status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
            })
          }
        }
        await _audit('createUser', 'Created user: ' + username.trim() + ' role: ' + safeRole)
        return new Response(JSON.stringify({ success: true, userId: uid }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        })
      }

      case 'deleteUser': {
        const { userId } = body
        if (!userId || typeof userId !== 'string' || !/^[0-9a-f-]{36}$/.test(userId)) {
          return new Response(JSON.stringify({ error: 'ID invalido' }), {
            status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        if (userId === adminUserId) {
          return new Response(JSON.stringify({ error: 'No puedes eliminarte a ti mismo' }), {
            status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        const { error: delError } = await adminClient.auth.admin.deleteUser(userId)
        if (delError) {
          return new Response(JSON.stringify({ error: 'Error al eliminar usuario' }), {
            status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        await adminClient.from('profiles').delete().eq('id', userId)
        await _audit('deleteUser', 'Deleted user ID: ' + userId)
        return new Response(JSON.stringify({ success: true }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        })
      }

      case 'changePassword': {
        const { userId: targetId, newPassword } = body
        if (!targetId || typeof targetId !== 'string' || !/^[0-9a-f-]{36}$/.test(targetId)) {
          return new Response(JSON.stringify({ error: 'ID invalido' }), {
            status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        if (!newPassword || typeof newPassword !== 'string' || newPassword.length < 8 || newPassword.length > 128 ||
            !/[A-Z]/.test(newPassword) || !/[a-z]/.test(newPassword) || !/[0-9]/.test(newPassword) ||
            !/[!@#$%^&*()_+\-=\[\]{};':"\\|,.<>\/?]/.test(newPassword)) {
          return new Response(JSON.stringify({ error: 'Contrasena: min 8, mayuscula, minuscula, numero y caracter especial' }), {
            status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        const { error: pwError } = await adminClient.auth.admin.updateUserById(targetId, { password: newPassword })
        if (pwError) {
          return new Response(JSON.stringify({ error: 'Error al cambiar contrasena' }), {
            status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        await _audit('changePassword', 'Changed password for user ID: ' + targetId)
        return new Response(JSON.stringify({ success: true }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        })
      }

      case 'listUsers': {
        const { data: profiles, error: listError } = await adminClient
          .from('profiles')
          .select('id, username, role, full_name, created_at')
          .order('created_at', { ascending: true })
        if (listError) {
          return new Response(JSON.stringify({ error: 'Error al listar usuarios' }), {
            status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        return new Response(JSON.stringify({ data: profiles || [] }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        })
      }

      default:
        return new Response(JSON.stringify({ error: 'Accion no valida' }), {
          status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        })
    }
  } catch (err) {
    return new Response(JSON.stringify({ error: 'Error interno: ' + (err instanceof Error ? err.message : String(err)) }), {
      status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    })
  }
})
