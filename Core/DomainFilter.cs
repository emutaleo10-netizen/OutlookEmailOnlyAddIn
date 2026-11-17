using System;
using System.Collections.Generic;
using OutlookEmailOnlyAddIn.Configuration;

namespace OutlookEmailOnlyAddIn.Core
{
    /// <summary>
    /// メールアドレスが除外対象（社内ローカルドメイン）かどうかを判定するクラス
    /// </summary>
    public static class DomainFilter
    {
        /// <summary>
        /// 指定されたメールアドレスがローカル（除外）ドメインに該当するか判定します。
        /// </summary>
        /// <param name="emailAddress">判定対象のメールアドレス</param>
        /// <param name="settings">現在のアドイン設定</param>
        /// <returns>除外対象ドメインの場合は true、書き換え対象の場合は false</returns>
        public static bool IsExcludedDomain(string emailAddress, AddInSettings settings)
        {
            if (string.IsNullOrWhiteSpace(emailAddress) || settings == null)
            {
                return false;
            }

            List<string> excludedDomains = settings.ExcludedDomains;
            if (excludedDomains == null || excludedDomains.Count == 0)
            {
                return false;
            }

            string domain = ExtractDomain(emailAddress);
            if (string.IsNullOrEmpty(domain))
            {
                return false;
            }

            foreach (string rawExcluded in excludedDomains)
            {
                if (string.IsNullOrWhiteSpace(rawExcluded))
                {
                    continue;
                }

                string targetDomain = rawExcluded.Trim().TrimStart('@').TrimStart('.');
                if (string.IsNullOrEmpty(targetDomain))
                {
                    continue;
                }

                // 完全一致（大文字小文字を区別しない）
                if (string.Equals(domain, targetDomain, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                // サブドメイン一致（例: user@sub.kinki-sharyo.local が kinki-sharyo.local にマッチ）
                if (settings.IncludeSubdomains &&
                    domain.EndsWith("." + targetDomain, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// メールアドレスから @ 以降のドメイン部分を抽出します。
        /// </summary>
        public static string ExtractDomain(string emailAddress)
        {
            if (string.IsNullOrWhiteSpace(emailAddress))
            {
                return null;
            }

            string trimmed = emailAddress.Trim().Trim('<', '>', '"', '\'');
            int atIndex = trimmed.LastIndexOf('@');
            if (atIndex <= 0 || atIndex >= trimmed.Length - 1)
            {
                return null;
            }

            return trimmed.Substring(atIndex + 1).Trim();
        }
    }
}
