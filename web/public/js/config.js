// Conexión con Supabase (inicio de sesión y base de datos en la nube).
// La clave «anon» es pública por diseño: los permisos los controla la base con
// seguridad a nivel de fila (RLS), según el rol de cada usuario en public.sgo_usuarios.
// Si se dejan vacíos, el programa funciona sin inicio de sesión y guarda solo en el navegador.
export const SUPABASE = {
  url: 'https://nlfxwxqymyttwbvmjomw.supabase.co',
  clave: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im5sZnh3eHF5bXl0dHdidm1qb213Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTM2NjUwNjgsImV4cCI6MjEwOTI0MTA2OH0.q7MD9FRf0SZfqBQvtiiBPscR8uYVAoyGtaIr9MChas0',
  funcionUsuarios: 'sgo-usuarios',
};
