-- ===========================================================================
-- Clientes del estudio — tabla en Supabase
-- ---------------------------------------------------------------------------
-- Pegá TODO esto en Supabase: panel del proyecto -> SQL Editor -> New query
-- -> pegar -> Run. Crea la tabla que usa la sección "Clientes".
-- ===========================================================================

create table if not exists public.clientes (
    id                     bigint generated always as identity primary key,
    razon_social           text not null,
    cuit                   text not null unique,
    tipo                   text not null default 'Monotributo',
        -- 'Monotributo' | 'Responsable Inscripto' | 'Exento' | 'Sociedad' | 'Otro'
    categoria_mono         text,                       -- 'A', 'B', ... 'K' (sólo monotributo)
    email                  text,
    telefono               text,
    domicilio              text,
    dia_cierre             int,                        -- día del mes (1-31)
    mes_cierre             int,                        -- mes (1-12), para sociedades
    observaciones          text,
    activo                 boolean not null default true,
    padron_data            jsonb,                      -- snapshot del padrón A5 (lo llena la consulta ARCA)
    padron_actualizado_en  timestamptz,                -- fecha/hora de la última consulta a ARCA
    creado_en              timestamptz not null default now()
);

create index if not exists idx_clientes_activo on public.clientes (activo);
create index if not exists idx_clientes_razon  on public.clientes (lower(razon_social));
