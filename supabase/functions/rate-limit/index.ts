import { createClient } from 'https://esm.sh/@supabase/supabase-js@2.49.1'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
}

const MAX_ATTEMPTS = 5
const LOCKOUT_MS = 60 * 1000
const CLEANUP_THRESHOLD = 100

Deno.serve(async (req: Request) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const supabaseAnonKey = Deno.env.get('SUPABASE_ANON_KEY')!
    const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!

    const body = await req.json()
    const { username, password } = body

    if (!username || !password || typeof username !== 'string' || typeof password !== 'string') {
      return new Response(JSON.stringify({ error: 'Datos incompletos' }), {
        status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    const trimUser = username.trim().toLowerCase()
    if (trimUser.length > 50 || password.length > 128) {
      return new Response(JSON.stringify({ error: 'Datos no validos' }), {
        status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    const ip = req.headers.get('x-forwarded-for')?.split(',')[0]?.trim() || req.headers.get('x-real-ip') || 'unknown'
    const key = trimUser + '::' + ip

    const adminClient = createClient(supabaseUrl, supabaseServiceKey, {
      auth: { persistSession: false }
    })

    // Check rate limit
    const { data: rlData } = await adminClient
      .from('rate_limits')
      .select('attempts, locked_until')
      .eq('key', key)
      .single()

    if (rlData) {
      if (rlData.locked_until) {
        const lockedUntil = new Date(rlData.locked_until).getTime()
        const now = Date.now()
        if (now < lockedUntil) {
          const rem = Math.ceil((lockedUntil - now) / 1000)
          return new Response(JSON.stringify({
            error: 'Demasiados intentos. Espera ' + rem + 's',
            locked: true,
            remaining: rem
          }), {
            status: 429, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          })
        }
        // Lockout expired, reset
        await adminClient.from('rate_limits').update({
          attempts: 0, locked_until: null, updated_at: new Date().toISOString()
        }).eq('key', key)
      }
    }

    // Attempt login via Supabase Auth
    const email = trimUser + '@debbiom.app'
    const anonClient = createClient(supabaseUrl, supabaseAnonKey, {
      auth: { persistSession: false }
    })

    const { data: authData, error: authError } = await anonClient.auth.signInWithPassword({
      email, password
    })

    if (authError) {
      // Login failed - increment attempts
      const currentAttempts = (rlData?.attempts || 0) + 1
      const lockUntil = currentAttempts >= MAX_ATTEMPTS
        ? new Date(Date.now() + LOCKOUT_MS).toISOString()
        : null

      await adminClient.from('rate_limits').upsert({
        key,
        attempts: currentAttempts,
        locked_until: lockUntil,
        updated_at: new Date().toISOString()
      }, { onConflict: 'key' })

      // Audit failed login
      await adminClient.from('audit_log').insert({
        admin_id: null,
        admin_username: trimUser,
        action: 'loginFail',
        detail: 'Attempt ' + currentAttempts + '/' + MAX_ATTEMPTS + (lockUntil ? ' LOCKED' : ''),
        ip,
        created_at: new Date().toISOString()
      }).catch(() => {})

      if (lockUntil) {
        return new Response(JSON.stringify({
          error: 'Cuenta bloqueada (60s)',
          locked: true,
          remaining: 60
        }), {
          status: 429, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        })
      }

      return new Response(JSON.stringify({
        error: 'Usuario o contrasena incorrectos (' + (MAX_ATTEMPTS - currentAttempts) + ' intentos)',
        attemptsLeft: MAX_ATTEMPTS - currentAttempts
      }), {
        status: 401, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    // Login successful - reset rate limit
    if (rlData) {
      await adminClient.from('rate_limits').update({
        attempts: 0, locked_until: null, updated_at: new Date().toISOString()
      }).eq('key', key)
    }

    // Return session to client
    return new Response(JSON.stringify({
      success: true,
      session: authData.session,
      user: authData.user
    }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    })

  } catch (err) {
    return new Response(JSON.stringify({ error: 'Error interno del servidor' }), {
      status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    })
  }
})
