        #region isWatchableFile
        /// <summary>
        /// ---inside bracket use
        /// @: Alpha Only [a-zA-Z]
        /// #: Numeric Only [0-9]
        /// +: AlphaNumeric [0-9a-zA-Z]
        /// .: Any character include special character
        /// [YYYYMMDD], [MMDDYYYY], [YYMMDD], [MMDDYY]
        /// 
        /// ---outside bracket use
        /// *: Any character and Any length
        /// |: or 
        /// (): group            (m|M|o)(c|C) same as [mMo][cC] but [] is not a group
        /// (|): group or group  (mc|oc|xx) --> 'mc' or 'oc' or 'xx'
        /// ---
        /// \/:*?"<>| is not allowed for filename
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="searchPattern"></param>
        /// <param name="ignoreCase"></param>
        /// <returns></returns>
        public static Boolean isWatchableFile(String fileName, String searchPattern, Boolean ignoreCase)
        {
            Boolean returnValue = false;
            string datePattern = string.Empty;
            try
            {
                // searchPattern : (mc|oc|xx)#.(adm|ADM|loc|LOC|svc|SVC)[YYYYMMDD]*.(txt|csv)
                Boolean hasDate = searchPattern.Contains("[");
                if (hasDate)
                {
                    datePattern = Regex.Match(searchPattern, @"\[(.+)\]").Groups[1].Value.ToLower();

                    string[] seperator = { "-", "/", "." };
                    bool hasSeperator = seperator.Any(x => datePattern.Contains(x));

                    if (hasSeperator)
                    {
                        if (datePattern.Length > 8)
                        {
                            if (datePattern.StartsWith("yyyy"))     //yyyy-mm-dd
                            {
                                searchPattern = searchPattern.ToLower().Replace("[yyyy-mm-dd]", "(19|20)[0-9][0-9][-/.](0[1-9]|1[012])[-/.](0[0-9]|[12][0-9]|3[01])");
                            }
                            else
                            {   //mm-dd-yyyy
                                searchPattern = searchPattern.ToLower().Replace("[mm-dd-yyyy]", "(0[1-9]|1[012])[-/.](0[0-9]|[12][0-9]|3[01])[-/.](19|20)[0-9][0-9]");
                            }
                        }
                        else
                        {
                            if (datePattern.StartsWith("yy"))     //yy-mm-dd
                            {
                                searchPattern = searchPattern.ToLower().Replace("[yy-mm-dd]", "[0-9][0-9][-/.](0[1-9]|1[012])[-/.](0[0-9]|[12][0-9]|3[01])");
                            }
                            else
                            {   //mm-dd-yy
                                searchPattern = searchPattern.ToLower().Replace("[mm-dd-yy]", "(0[1-9]|1[012])[-/.](0[0-9]|[12][0-9]|3[01])[-/.][0-9][0-9]");
                            }
                        }
                    }
                    else
                    {
                        if (datePattern.Length > 6)
                        {
                            if (datePattern.StartsWith("yyyy"))     //yyyymmdd
                            {
                                searchPattern = searchPattern.ToLower().Replace("[yyyymmdd]", "(19|20)[0-9][0-9](0[1-9]|1[012])(0[0-9]|[12][0-9]|3[01])");
                            }
                            else
                            {   //mmddyyyy
                                searchPattern = searchPattern.ToLower().Replace("[mmddyyyy]", "(0[1-9]|1[012])(0[0-9]|[12][0-9]|3[01])(19|20)[0-9][0-9]");
                            }
                        }
                        else
                        {
                            if (datePattern.StartsWith("yy"))     //yymmdd
                            {
                                searchPattern = searchPattern.ToLower().Replace("[yymmdd]", "[0-9][0-9](0[1-9]|1[012])(0[0-9]|[12][0-9]|3[01])");
                            }
                            else
                            {   //mmddyy
                                searchPattern = searchPattern.ToLower().Replace("[mmddyy]", "(0[1-9]|1[012])(0[0-9]|[12][0-9]|3[01])[0-9][0-9]");
                            }
                        }
                    }

                }

                //translate
                searchPattern = searchPattern.Replace("@", "[a-zA-Z]");
                searchPattern = searchPattern.Replace("#", "[0-9]");
                searchPattern = searchPattern.Replace("+", "[0-9a-zA-Z]");
                //searchPattern = searchPattern.Replace(".", "Any character (except \n newline)");

                //change '*' to '.*' for Regex use
                searchPattern = searchPattern.Replace("*", ".*");


                //Regex.Matches
                MatchCollection matches;
                if (ignoreCase)
                {
                    matches = Regex.Matches(fileName, searchPattern, RegexOptions.IgnoreCase);
                    foreach (Match match in matches)
                    {
                        foreach (Capture capture in match.Captures)
                        {
                            Console.WriteLine("Index={0}, Value={1}", capture.Index, capture.Value);
                        }
                    }
                }
                else
                {
                    matches = Regex.Matches(fileName, searchPattern);
                    foreach (Match match in matches)
                    {
                        foreach (Capture capture in match.Captures)
                        {
                            Console.WriteLine("Index={0}, Value={1}", capture.Index, capture.Value);
                        }
                    }
                }

                //check Match results
                if (matches.Count > 0)
                    returnValue = matches[0].Groups[0].Captures[0].Length == fileName.Length ? true : false;
                else
                    returnValue = false;


                return returnValue;
            }
            catch (Exception ex)
            {
                String message = String.Format(" Error checking file {1} with pattern {2}, Message={0}", ex.Message, fileName, searchPattern);
                if (ex.StackTrace != null)
                    message += String.Format(", {0}", ex.StackTrace);

                throw new Exception(message);
            }
        }
        #endregion
