# Supabase接続設定

大会データは Supabase プロジェクト `bond sift` の `public.football_tournaments` に保存します。

本番環境には次の環境変数を設定してください。

```text
SUPABASE_URL=https://xvycsjiqtwthdihithnq.supabase.co
SUPABASE_SECRET_KEY=SupabaseのSecret key
```

`SUPABASE_SECRET_KEY` は Supabase Dashboard の「Project Settings → API Keys」から取得し、Renderなどのサーバー環境変数にだけ設定してください。ブラウザへ公開される `NEXT_PUBLIC_` 変数には設定しないでください。

テーブルはRLSを有効化し、`anon` と `authenticated` からの直接アクセスを無効にしています。読み書きは `/api/tournaments` を通してサーバー側だけで実行します。
