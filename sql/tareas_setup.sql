-- ===========================================================================
-- Tareas y checklist — tablas en Supabase
-- ---------------------------------------------------------------------------
-- Pegá TODO esto en Supabase: panel del proyecto -> SQL Editor -> New query
-- -> pegar -> Run. Crea las 3 tablas que usa la sección "Tareas y checklist".
-- ===========================================================================

-- Tareas puntuales (de una vez)
create table if not exists public.tareas (
    id          bigint generated always as identity primary key,
    titulo      text not null,
    nota        text,
    monto       numeric,
    creada_por  text,
    creada_en   timestamptz not null default now(),
    hecha       boolean not null default false,
    hecha_por   text,
    hecha_en    timestamptz
);

-- Definición de las tareas rutinarias (se repiten)
create table if not exists public.rutinas (
    id          bigint generated always as identity primary key,
    titulo      text not null,
    frecuencia  text not null default 'diaria',   -- 'diaria' | 'semanal' | 'mensual'
    orden       int  not null default 0,
    activa      boolean not null default true,
    creada_en   timestamptz not null default now()
);

-- Marcas de completado de cada rutina, por período
-- (que exista una fila = esa rutina está hecha en ese período)
create table if not exists public.rutinas_estado (
    id          bigint generated always as identity primary key,
    rutina_id   bigint not null references public.rutinas(id) on delete cascade,
    periodo     text not null,            -- '2026-06-19' | '2026-W25' | '2026-06'
    hecha_por   text,
    hecha_en    timestamptz not null default now(),
    unique (rutina_id, periodo)
);
