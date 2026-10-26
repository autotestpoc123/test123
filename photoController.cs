using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Net.Http.Headers;
using COD.FirmwideDirectory.API.Models.Options;
using COD.FirmwideDirectory.API.Services;

namespace COD.FirmwideDirectory.API.Controllers
{
    [Authorize]
    [RequiredScope("api")]
    [ApiVersion("1.0")]
    [ApiExplorerSettings(GroupName = "v1.0")]
    [Route("api/v{version:apiVersion}/photos")]
    [ApiController]
    public class PhotoController : ControllerBase
    {
        private readonly ILogger<PhotoController> _logger;

        private readonly PhotoOptions _photoOptions;

        private readonly IPhotoThumbnailCache _thumbnailCache;

        public PhotoController(
            ILogger<PhotoController> logger,
            IOptions<PhotoOptions> photoOptions,
            IPhotoThumbnailCache thumbnailCache)
        {
            _logger = logger;
            _photoOptions = photoOptions.Value;
            _thumbnailCache = thumbnailCache;
        }
        #region API
        /// <summary>
        /// Get photo by msid.
        /// Without size, returns the original file.
        /// With size, returns a JPEG resized to the nearest allowed width, served from the in-memory thumbnail cache.
        /// </summary>
        /// <param name="msid"></param>
        /// <param name="size">Requested width in pixels. Normalized to PhotoOptions.ThumbnailWidths. Omit for original.</param>
        /// <returns></returns>
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status401Unauthorized)]
        [ProducesResponseType(StatusCodes.Status403Forbidden)]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status304NotModified)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        [Route("{msid}")]
        [HttpGet]
        public IActionResult GetPhoto(string msid, [FromQuery] int? size = null)
        {
            if (string.IsNullOrWhiteSpace(msid))
            {
                return BadRequest("key is empty");
            }

            var msidTrimmed = msid.Trim();
            if (msidTrimmed.Length < ControllerConstants.MSIDMinLength)
            {
                return BadRequest($"Key length must be at least {ControllerConstants.MSIDMinLength}");
            }

            if (msidTrimmed.Length > ControllerConstants.MSIDMaxLength)
            {
                return BadRequest($"Key can not be longer than {ControllerConstants.MSIDMaxLength}");
            }

            if (Utility.IsValidMSIDForPhoto(msidTrimmed) == false)
            {
                return BadRequest("Invalid MSID");
            }

            int? thumbnailWidth = null;
            if (size.HasValue)
            {
                thumbnailWidth = _photoOptions.NormalizeThumbnailWidth(size.Value);
                if (thumbnailWidth == null)
                {
                    return BadRequest("Invalid photo size");
                }
            }
            string contentType;
            if (_photoOptions.PhotoType == ".jpg")
            {
                contentType = "image/jpeg";
            }
            else
            {
                return StatusCode(StatusCodes.Status500InternalServerError,
                    new { message = $"Unsupported file type: {_photoOptions.PhotoType}" });
            }

            var photoFullPath = Utility.GetUserPhotoFullPath(msidTrimmed, _photoOptions);
            var fileInfo = new FileInfo(photoFullPath);
            if (fileInfo.Exists == false)
            {
                // user might not have photo yet. Return 204 instead of 404
                return NoContent();
            }

            var lastModified = fileInfo.LastWriteTimeUtc;
            // https://developer.mozilla.org/en-US/docs/Web/HTTP/Reference/Headers/Last-Modified
            var lastModifiedGMT = lastModified.ToString("R");
            var cacheControl = $"private, max-age={_photoOptions.PhotoCacheControl}";

            // https://developer.mozilla.org/en-US/docs/Web/HTTP/Reference/Status/304
            // Thumbnail identity includes the normalized width so a small and large render do not share
            // an ETag, and the encoding-rules version so quality/codec changes invalidate browser caches.
            string rawMetaData = msidTrimmed + lastModified.Ticks.ToString() + fileInfo.Length
                + (thumbnailWidth.HasValue ? ":" + thumbnailWidth.Value + ":" + _photoOptions.ThumbnailVersion : "");
            var entityTag = new EntityTagHeaderValue("\"" + Utility.GenerateETag(rawMetaData) + "\"");
            string eTag = entityTag.ToString();
            var hasIfNoneMatch = Request.Headers.ContainsKey(HeaderNames.IfNoneMatch);
            if (hasIfNoneMatch
                && EntityTagHeaderValue.TryParseList(Request.Headers[HeaderNames.IfNoneMatch].Select(value => value ?? string.Empty).ToArray(), out var requestTags)
                && requestTags.Any(tag => tag == EntityTagHeaderValue.Any || tag.Compare(entityTag, useStrongComparison: false)))
            {
                SetCacheHeaders(eTag, cacheControl, lastModifiedGMT);
                return StatusCode(StatusCodes.Status304NotModified);
            }

            // Source mtime alone cannot validate a thumbnail after encoder/version changes.
            if (!hasIfNoneMatch && !thumbnailWidth.HasValue && Request.Headers.ContainsKey(HeaderNames.IfModifiedSince))
            {
                if (DateTime.TryParse(Request.Headers.IfModifiedSince, out var ifModifiedSince))
                {
                    if (lastModified <= ifModifiedSince.ToUniversalTime())
                    {
                        SetCacheHeaders(eTag, cacheControl, lastModifiedGMT);
                        return StatusCode(StatusCodes.Status304NotModified);
                    }
                }
            }

            if (thumbnailWidth.HasValue)
            {
                // Keep pending generation non-cacheable. Error middleware must preserve
                // this policy (or set no-store itself) if it clears response headers.
                Response.Headers.CacheControl = "no-store";
                var bytes = _thumbnailCache.GetOrCreate(photoFullPath, lastModified.Ticks, thumbnailWidth.Value);
                if (bytes == null)
                {
                    _logger.LogWarning("Thumbnail generation failed for {Msid} width {Width}", msidTrimmed, thumbnailWidth.Value);
                    return StatusCode(StatusCodes.Status500InternalServerError,
                        new { message = "Photo could not be resized" });
                }
                SetCacheHeaders(eTag, cacheControl, lastModifiedGMT);
                return File(bytes, contentType);
            }

            SetCacheHeaders(eTag, cacheControl, lastModifiedGMT);
            return PhysicalFile(photoFullPath, contentType);
        }

        private void SetCacheHeaders(string eTag, string cacheControl, string lastModifiedGmt)
        {
            Response.Headers.ETag = eTag;
            Response.Headers.CacheControl = cacheControl;
            Response.Headers.LastModified = lastModifiedGmt;
        }

        #endregion
    }
}
