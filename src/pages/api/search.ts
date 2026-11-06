import { type NextRequest, NextResponse } from 'next/server'
import axios from 'redaxios'
import apiConfig from '~config/api.config'
import siteConfig from '~config/site.config'
import { getAccessToken } from '.'

export const runtime = 'edge'

/**
 * Why this no longer uses the Microsoft Graph Search API (POST /v1.0/search/query):
 *
 * - That endpoint is backed by the Substrate search service, which is unreliable
 *   for personal OneDrive accounts: it frequently fails with 502 Bad Gateway /
 *   500 ("Error Calling Substrate Search") or returns stale and empty results,
 *   even when the same files are found fine in the OneDrive web UI.
 * - Microsoft's own support answer to this exact problem recommends the
 *   OneDrive drive search endpoint instead, and the old code passed the
 *   upstream 502 straight through to the browser, which is the "Error:
 *   Request failed with status code 502" the search box shows.
 *
 * The drive search endpoint (GET {driveApi}/root/search) runs against the
 * drive itself, works for personal and business accounts, and needs only the
 * Files.Read.All scope this app already requests. It returns full driveItem
 * objects; we still return just the ids to keep the frontend contract
 * (/api/search -> string[], then /api/item per id) unchanged.
 */

/**
 * Escape the search query for the `q` parameter of the drive search endpoint.
 * The query is wrapped in single quotes in the URL, so a literal single quote
 * inside it is doubled, per OData string escaping. URL-encoding happens at
 * the call site.
 */
function sanitiseQuery(query: string): string {
  return query.replace(/'/g, "''")
}

export default async function handler(req: NextRequest): Promise<Response> {
  // Query parameter from request
  const { q: searchQuery = '' } = Object.fromEntries(req.nextUrl.searchParams)
  const query = searchQuery.trim()

  if (!query) {
    return new NextResponse(JSON.stringify([]), {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': apiConfig.cacheControlHeader,
      },
    })
  }

  // Get access token from storage. This used to sit outside the try block, so
  // any failure reading the token (for example a missing or misconfigured KV
  // binding) crashed the edge function, which the hosting platform surfaces
  // as a 502. It now fails closed with a JSON error instead.
  let accessToken: string
  try {
    accessToken = await getAccessToken()
  } catch {
    return new Response(JSON.stringify({ error: 'Failed to read access token.' }), { status: 500 })
  }

  if (!accessToken) {
    return new Response(JSON.stringify({ error: 'No access token.' }), { status: 403 })
  }

  try {
    const { data } = await axios.get(`${apiConfig.driveApi}/root/search(q='${encodeURIComponent(sanitiseQuery(query))}')`, {
      headers: {
        // The old code sent the raw token with no "Bearer " scheme, which
        // Microsoft Graph rejects with 401.
        Authorization: `Bearer ${accessToken}`,
      },
      params: {
        $top: siteConfig.maxItems,
        $select: 'id',
      },
    })

    const itemIds = (data.value ?? []).map((item: { id: string }) => item.id)

    return new NextResponse(JSON.stringify(itemIds), {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': apiConfig.cacheControlHeader,
      },
    })
  } catch (error: any) {
    return new Response(JSON.stringify({ error: error?.response?.data ?? 'Internal server error.' }), {
      status: error?.response?.status ?? 500,
    })
  }
}
