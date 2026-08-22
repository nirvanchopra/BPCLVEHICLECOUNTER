import { useState } from 'react'
import { supabase } from './supabaseClient'

function Login() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')

  const handleLogin = async (e) => {
    e.preventDefault()
    setError('')

    const { error } = await supabase.auth.signInWithPassword({
      email,
      password
    })

    if (error) {
      setError(error.message)
    }
  }

  return (
    <div style={{
      padding: 20,
      maxWidth: 400,
      margin: '40px auto'
    }}>
      <h2>BPCL Vehicle Counter</h2>

      <form onSubmit={handleLogin}>
        <input
          type="email"
          placeholder="Email"
          value={email}
          onChange={e => setEmail(e.target.value)}
          style={{
            width: '100%',
            padding: 12,
            marginBottom: 10,
            fontSize: 16
          }}
          required
        />

        <input
          type="password"
          placeholder="Password"
          value={password}
          onChange={e => setPassword(e.target.value)}
          style={{
            width: '100%',
            padding: 12,
            marginBottom: 10,
            fontSize: 16
          }}
          required
        />

        {error && (
          <p style={{ color: 'red' }}>
            {error}
          </p>
        )}

        <button
          type="submit"
          style={{
            width: '100%',
            padding: 14,
            fontSize: 16
          }}
        >
          Sign In
        </button>
      </form>
    </div>
  )
}

export default Login